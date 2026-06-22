#!/usr/bin/env python3
import argparse
import asyncio
import contextlib
import logging
from dataclasses import dataclass
from itertools import cycle
from typing import Dict, List, Optional
from urllib.parse import urlsplit

from aiohttp import ClientSession, ClientTimeout, WSMsgType, web


LOG = logging.getLogger("load_balancer")
STICKY_COOKIE = "lb_backend"
HOP_BY_HOP_HEADERS = {
    "connection",
    "keep-alive",
    "proxy-authenticate",
    "proxy-authorization",
    "te",
    "trailers",
    "transfer-encoding",
    "upgrade",
}


@dataclass
class Backend:
    base_url: str
    healthy: bool = True

    @property
    def host_port(self) -> str:
        parsed = urlsplit(self.base_url)
        return parsed.netloc


class BackendPool:
    def __init__(self, backends: List[str]):
        if not backends:
            raise ValueError("At least one backend is required")
        self._backends: Dict[str, Backend] = {url: Backend(base_url=url) for url in backends}
        self._cycle = cycle(list(self._backends.keys()))
        self._lock = asyncio.Lock()

    def mark_health(self, base_url: str, healthy: bool):
        backend = self._backends.get(base_url)
        if backend:
            backend.healthy = healthy

    def get_by_host_port(self, host_port: str) -> Optional[str]:
        for backend in self._backends.values():
            if backend.host_port == host_port:
                return backend.base_url
        return None

    def healthy_backends(self) -> List[str]:
        return [b.base_url for b in self._backends.values() if b.healthy]

    async def choose_backend(self, sticky_host_port: Optional[str]) -> Optional[str]:
        async with self._lock:
            healthy = self.healthy_backends()
            if not healthy:
                return None

            if sticky_host_port:
                sticky_backend = self.get_by_host_port(sticky_host_port)
                if sticky_backend in healthy:
                    return sticky_backend

            for _ in range(len(self._backends)):
                candidate = next(self._cycle)
                if candidate in healthy:
                    return candidate
            return healthy[0]


def _filtered_headers(headers) -> Dict[str, str]:
    result = {}
    for key, value in headers.items():
        if key.lower() not in HOP_BY_HOP_HEADERS and key.lower() != "host":
            result[key] = value
    return result


async def _proxy_websocket(request: web.Request, target_base: str, client: ClientSession) -> web.StreamResponse:
    subprotocol_header = request.headers.get("Sec-WebSocket-Protocol", "")
    subprotocols = [proto.strip() for proto in subprotocol_header.split(",") if proto.strip()]

    client_ws = web.WebSocketResponse(protocols=subprotocols)
    await client_ws.prepare(request)

    target_url = target_base.replace("http://", "ws://", 1).replace("https://", "wss://", 1)
    target_url = f"{target_url}{request.rel_url}"
    backend_headers = _filtered_headers(request.headers)

    async with client.ws_connect(target_url, headers=backend_headers, protocols=subprotocols) as server_ws:
        async def client_to_server():
            async for msg in client_ws:
                if msg.type == WSMsgType.TEXT:
                    await server_ws.send_str(msg.data)
                elif msg.type == WSMsgType.BINARY:
                    await server_ws.send_bytes(msg.data)
                elif msg.type == WSMsgType.CLOSE:
                    await server_ws.close()

        async def server_to_client():
            async for msg in server_ws:
                if msg.type == WSMsgType.TEXT:
                    await client_ws.send_str(msg.data)
                elif msg.type == WSMsgType.BINARY:
                    await client_ws.send_bytes(msg.data)
                elif msg.type == WSMsgType.CLOSE:
                    await client_ws.close()

        await asyncio.gather(client_to_server(), server_to_client())

    return client_ws


async def _proxy_http(request: web.Request, target_base: str, client: ClientSession) -> web.StreamResponse:
    target_url = f"{target_base}{request.rel_url}"
    body = await request.read()
    headers = _filtered_headers(request.headers)

    async with client.request(
        method=request.method,
        url=target_url,
        headers=headers,
        data=body,
        allow_redirects=False,
    ) as upstream:
        response_headers = _filtered_headers(upstream.headers)
        response = web.Response(
            status=upstream.status,
            headers=response_headers,
            body=await upstream.read(),
        )
        return response


async def _health_monitor(app: web.Application):
    pool: BackendPool = app["pool"]
    client: ClientSession = app["client"]
    interval: int = app["health_interval"]
    timeout = ClientTimeout(total=2)

    while True:
        for backend in list(pool._backends.keys()):
            health_url = f"{backend}/_stcore/health"
            backend_state = pool._backends[backend]
            was_healthy = backend_state.healthy
            healthy = False
            try:
                async with client.get(health_url, timeout=timeout) as resp:
                    healthy = resp.status < 500
            except Exception:
                healthy = False

            pool.mark_health(backend, healthy)
            if healthy != was_healthy:
                state = "healthy" if healthy else "unhealthy"
                LOG.info("Backend %s is now %s", backend, state)

        await asyncio.sleep(interval)


async def _request_handler(request: web.Request) -> web.StreamResponse:
    pool: BackendPool = request.app["pool"]
    client: ClientSession = request.app["client"]
    sticky_cookie = request.cookies.get(STICKY_COOKIE)
    target_base = await pool.choose_backend(sticky_cookie)

    if not target_base:
        return web.Response(status=503, text="No healthy backend is available")

    connection_header = request.headers.get("Connection", "")
    upgrade_header = request.headers.get("Upgrade", "")
    is_ws = "upgrade" in connection_header.lower() and upgrade_header.lower() == "websocket"

    try:
        if is_ws:
            return await _proxy_websocket(request, target_base, client)

        response = await _proxy_http(request, target_base, client)
        response.set_cookie(
            STICKY_COOKIE,
            value=urlsplit(target_base).netloc,
            max_age=86400,
            httponly=True,
            samesite="Lax",
        )
        return response
    except Exception:
        LOG.exception("Proxy request failed for backend %s", target_base)
        return web.Response(status=502, text="Upstream request failed")


async def _on_startup(app: web.Application):
    app["client"] = ClientSession(timeout=ClientTimeout(total=30))
    app["monitor_task"] = asyncio.create_task(_health_monitor(app))


async def _on_cleanup(app: web.Application):
    task = app.get("monitor_task")
    if task:
        task.cancel()
        with contextlib.suppress(asyncio.CancelledError):
            await task
    await app["client"].close()


def _build_app(backends: List[str], health_interval: int) -> web.Application:
    app = web.Application()
    app["pool"] = BackendPool(backends)
    app["health_interval"] = health_interval
    app.router.add_route("*", "/{tail:.*}", _request_handler)
    app.on_startup.append(_on_startup)
    app.on_cleanup.append(_on_cleanup)
    return app


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Round-robin load balancer for Streamlit replicas")
    parser.add_argument("--listen-host", default="127.0.0.1")
    parser.add_argument("--listen-port", type=int, default=8501)
    parser.add_argument("--backends", required=True, help="Comma-separated backend URLs")
    parser.add_argument("--health-interval", type=int, default=5)
    return parser.parse_args()


def main():
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
    args = parse_args()
    backends = [b.strip().rstrip("/") for b in args.backends.split(",") if b.strip()]
    app = _build_app(backends=backends, health_interval=args.health_interval)
    web.run_app(app, host=args.listen_host, port=args.listen_port)


if __name__ == "__main__":
    main()