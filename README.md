# MedComms Reference QA and Fact-Checking Workbench

This repository contains a Streamlit-based scientific reference QA application focused on:
- source attribution
- citation verification
- local full-text evidence search
- copyright and rights screening
- reviewer-oriented reporting

## Current App
The application entry point is `streamlit_app.py`.

Run locally:

1. Install dependencies

   ```bash
   pip install -r requirements.txt
   ```

2. Start the app

   ```bash
   streamlit run streamlit_app.py
   ```

## Load-Balanced Run Mode
The repository now supports local load balancing for Streamlit replicas using `load_balancer.py`.

Run with the existing tunnel script (load balancing enabled by default):

```bash
./refresh_public_tunnel.sh
```

Optional environment variables:

- `ENABLE_LOAD_BALANCING=1` enables the balancer (`0` keeps single-instance mode).
- `STREAMLIT_REPLICAS=2` sets number of Streamlit replicas.
- `STREAMLIT_BASE_PORT=8601` sets first replica port.
- `TUNNEL_PROVIDER=cloudflared` or `TUNNEL_PROVIDER=localhostrun` selects tunnel backend.

In load-balanced mode:

- Streamlit replicas run on `STREAMLIT_BASE_PORT` and upward.
- `load_balancer.py` binds to port `8501` and routes requests across healthy replicas.
- Sticky cookies are used to keep browser sessions on one backend when possible.

## Future Enhancement Roadmap
The future-state platform requirements and phased implementation plan are documented in:

- `FUTURE_ENHANCEMENT_VERIFICATION_FRAMEWORK.md`

This roadmap defines the transition from literature search behavior to an evidence verification framework with:
- prioritized full-text retrieval
- section-level evidence grounding
- citation context validation
- NLI-based support/contradiction scoring
- semantic vector retrieval
- rights-aware processing
- mandatory human-review transparency

<table>
<tr>
<td align="center">
<div style="width:50px;height:50px;background:#12344D;border:1px solid #ccc;"></div>
<br>Primary
<br>#12344D
</td>

<td align="center">
<div style="width:50px;height:50px;background:#1F4E79;border:1px solid #ccc;"></div>
<br>Secondary
<br>#1F4E79
</td>

<td align="center">
<div style="width:50px;height:50px;background:#2F80C1;border:1px solid #ccc;"></div>
<br>Accent
<br>#2F80C1
</td>
</tr>
</table>
