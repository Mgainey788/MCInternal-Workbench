#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="/workspaces/medcomms_reference_qa_staff_workbench_simplified_fulltext_REVISED_ALL_LIT_SCREENING.py"
APP_PORT="8501"
STREAMLIT_REPLICAS="${STREAMLIT_REPLICAS:-2}"
STREAMLIT_BASE_PORT="${STREAMLIT_BASE_PORT:-8601}"
ENABLE_LOAD_BALANCING="${ENABLE_LOAD_BALANCING:-1}"
TUNNEL_KEY="${HOME}/.ssh/localhostrun_ed25519"
MAX_WAIT_SECONDS="60"
TUNNEL_PROVIDER="${TUNNEL_PROVIDER:-cloudflared}"


STREAMLIT_MATCH='(streamlit run streamlit_app.py|python(3)? -m streamlit run streamlit_app.py)'
VENV_CANDIDATES=(".venv-1" ".venv")
VENV_DIR=""

for candidate in "${VENV_CANDIDATES[@]}"; do
  if [[ -f "${ROOT_DIR}/${candidate}/bin/activate" ]]; then
    VENV_DIR="${ROOT_DIR}/${candidate}"
    break
  fi
done

if [[ -z "$VENV_DIR" ]]; then
  echo "No supported virtual environment found. Expected one of: ${VENV_CANDIDATES[*]}"
  exit 1
fi

echo "[1/4] Stopping existing Streamlit and tunnel processes..."
pkill -f "$STREAMLIT_MATCH" || true
pkill -f "python(3)? .*load_balancer.py" || true
pkill -f "ssh .*localhost.run|ssh.localhost.run" || true
pkill -f "cloudflared tunnel --no-autoupdate --url" || true

# Ensure prior Streamlit processes have exited before starting a fresh instance.
for _ in $(seq 1 10); do
  if ! pgrep -f "$STREAMLIT_MATCH" >/dev/null 2>&1; then
    break
  fi
  sleep 1
done

echo "[2/4] Starting Streamlit services..."
cd "$ROOT_DIR"
source "${VENV_DIR}/bin/activate"

declare -a BACKEND_PORTS=()
FRONTEND_LOG_HINT="/tmp/streamlit.log"
if [[ "$ENABLE_LOAD_BALANCING" == "1" ]]; then
  FRONTEND_LOG_HINT="/tmp/load_balancer.log and /tmp/streamlit_*.log"
  if [[ "$STREAMLIT_REPLICAS" -lt 2 ]]; then
    STREAMLIT_REPLICAS="2"
  fi
  for idx in $(seq 0 $((STREAMLIT_REPLICAS - 1))); do
    port=$((STREAMLIT_BASE_PORT + idx))
    BACKEND_PORTS+=("$port")
    nohup python -m streamlit run streamlit_app.py \
      --server.port "$port" \
      --server.enableCORS false \
      --server.enableXsrfProtection false > "/tmp/streamlit_${port}.log" 2>&1 &
  done
else
  BACKEND_PORTS+=("$APP_PORT")
  nohup python -m streamlit run streamlit_app.py \
    --server.port "$APP_PORT" \
    --server.enableCORS false \
    --server.enableXsrfProtection false > /tmp/streamlit.log 2>&1 &
fi

echo "[3/4] Waiting for local app health..."
for backend_port in "${BACKEND_PORTS[@]}"; do
  healthy="0"
  for _ in $(seq 1 "$MAX_WAIT_SECONDS"); do
    if ss -ltn | grep -q ":${backend_port} "; then
      if curl -sS --max-time 2 -o /dev/null -I "http://127.0.0.1:${backend_port}" 2>/dev/null; then
        healthy="1"
        break
      fi
    fi
    sleep 1
  done

  if [[ "$healthy" != "1" ]]; then
    backend_log="/tmp/streamlit_${backend_port}.log"
    if [[ "$ENABLE_LOAD_BALANCING" != "1" ]]; then
      backend_log="/tmp/streamlit.log"
    fi
    echo "Streamlit did not become healthy on port ${backend_port}."
    echo "Check logs: ${backend_log}"
    echo "Last 40 log lines:"
    tail -n 40 "${backend_log}" || true
    echo "Active streamlit processes:"
    pgrep -af "$STREAMLIT_MATCH" || true
    exit 1
  fi
done

if [[ "$ENABLE_LOAD_BALANCING" == "1" ]]; then
  backends_csv=""
  for backend_port in "${BACKEND_PORTS[@]}"; do
    if [[ -z "$backends_csv" ]]; then
      backends_csv="http://127.0.0.1:${backend_port}"
    else
      backends_csv="${backends_csv},http://127.0.0.1:${backend_port}"
    fi
  done

  nohup python load_balancer.py \
    --listen-host 127.0.0.1 \
    --listen-port "$APP_PORT" \
    --backends "$backends_csv" \
    --health-interval 5 > /tmp/load_balancer.log 2>&1 &

  for _ in $(seq 1 "$MAX_WAIT_SECONDS"); do
    if ss -ltn | grep -q ":${APP_PORT} "; then
      if curl -sS --max-time 2 -o /dev/null -I "http://127.0.0.1:${APP_PORT}" 2>/dev/null; then
        break
      fi
    fi
    sleep 1
  done

  if ! ss -ltn | grep -q ":${APP_PORT} "; then
    echo "Load balancer did not bind to port ${APP_PORT}."
    echo "Check logs: /tmp/load_balancer.log"
    tail -n 40 /tmp/load_balancer.log || true
    exit 1
  fi

  if ! curl -sS --max-time 2 -o /dev/null -I "http://127.0.0.1:${APP_PORT}"; then
    echo "Load balancer is listening on ${APP_PORT} but health probe failed."
    echo "Check logs: /tmp/load_balancer.log"
    tail -n 40 /tmp/load_balancer.log || true
    exit 1
  fi
fi

if [[ "$ENABLE_LOAD_BALANCING" != "1" ]]; then
  if ! ss -ltn | grep -q ":${APP_PORT} "; then
    echo "Streamlit did not bind to port ${APP_PORT}."
    echo "Check logs: /tmp/streamlit.log"
    echo "Last 40 log lines:"
    tail -n 40 /tmp/streamlit.log || true
    echo "Active streamlit processes:"
    pgrep -af "$STREAMLIT_MATCH" || true
    exit 1
  fi

  if ! curl -sS --max-time 2 -o /dev/null -I "http://127.0.0.1:${APP_PORT}"; then
    echo "Streamlit is listening on ${APP_PORT} but health probe failed."
    echo "Check logs: /tmp/streamlit.log"
    echo "Last 40 log lines:"
    tail -n 40 /tmp/streamlit.log || true
    echo "Active streamlit processes:"
    pgrep -af "$STREAMLIT_MATCH" || true
    exit 1
  fi
fi

for _ in $(seq 1 "$MAX_WAIT_SECONDS"); do
  if ss -ltn | grep -q ":${APP_PORT} "; then
    if curl -sS --max-time 2 -o /dev/null -I "http://127.0.0.1:${APP_PORT}" 2>/dev/null; then
      break
    fi
  fi
  sleep 1
done

if ! ss -ltn | grep -q ":${APP_PORT} "; then
  echo "Application frontend did not bind to port ${APP_PORT}."
  echo "Check logs: ${FRONTEND_LOG_HINT}"
  echo "Last 40 log lines:"
  if [[ "$ENABLE_LOAD_BALANCING" == "1" ]]; then
    tail -n 40 /tmp/load_balancer.log || true
  else
    tail -n 40 /tmp/streamlit.log || true
  fi
  echo "Active streamlit processes:"
  pgrep -af "$STREAMLIT_MATCH" || true
  exit 1
fi

if ! curl -sS --max-time 2 -o /dev/null -I "http://127.0.0.1:${APP_PORT}"; then
  echo "Streamlit is listening on ${APP_PORT} but health probe failed."
  echo "Check logs: ${FRONTEND_LOG_HINT}"
  echo "Last 40 log lines:"
  if [[ "$ENABLE_LOAD_BALANCING" == "1" ]]; then
    tail -n 40 /tmp/load_balancer.log || true
  else
    tail -n 40 /tmp/streamlit.log || true
  fi
  echo "Active streamlit processes:"
  pgrep -af "$STREAMLIT_MATCH" || true
  exit 1
fi

echo "[4/4] Starting public tunnel..."
if [[ "$TUNNEL_PROVIDER" == "cloudflared" ]]; then
  if [[ ! -x "${ROOT_DIR}/cloudflared" ]]; then
    echo "cloudflared binary not found or not executable at ${ROOT_DIR}/cloudflared"
    echo "Set TUNNEL_PROVIDER=localhostrun to use SSH tunnel instead."
    exit 1
  fi

  exec "${ROOT_DIR}/cloudflared" tunnel --no-autoupdate --url "http://127.0.0.1:${APP_PORT}"
fi

if [[ -f "$TUNNEL_KEY" ]]; then
  exec ssh -i "$TUNNEL_KEY" -o IdentitiesOnly=yes -o ServerAliveInterval=60 -o ServerAliveCountMax=3 -R 80:localhost:${APP_PORT} ssh.localhost.run
else
  echo "SSH key not found at $TUNNEL_KEY. Falling back to nokey mode."
  exec ssh -o StrictHostKeyChecking=no -o ServerAliveInterval=60 -R 80:localhost:${APP_PORT} nokey@localhost.run
fi
