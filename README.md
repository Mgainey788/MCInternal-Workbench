# MedComms Reference QA and Fact-Checking Workbench

This repository contains a Streamlit-based scientific reference QA application focused on:
- source attribution
- citation verification
- local full-text evidence search
- copyright and rights screening
- reviewer-oriented reporting

## ✏️ How to Edit and Deploy

**One rule:** only edit `streamlit_app.py`. That is the file both Azure and Streamlit Cloud deploy.

### Edit → Deploy Workflow

```
1. Edit  streamlit_app.py
2. Commit your changes
3. Push / merge to the main branch
4. GitHub Actions auto-deploys to Azure (~2-3 min)
5. For Streamlit Cloud: confirm it is also pointed at main → streamlit_app.py
```

### Verify a deployment landed
After a push to `main`, go to **GitHub → Actions** and open the latest workflow run.  
The final step prints: `Deployed commit <SHA> to <app-name>`.  
Compare that SHA against the commit you just pushed. If they match, your changes are live.  
If the Azure UI still shows old content, do a hard-refresh (`Ctrl+Shift+R`) or restart the App Service.

### Manual deploy
The deployment workflow supports **workflow_dispatch** — go to **GitHub → Actions → Build and deploy Python app to Azure Web App - MCInternalWorkbench → Run workflow** to trigger a deploy without pushing new code.

---

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
