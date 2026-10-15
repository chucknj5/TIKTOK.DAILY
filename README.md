# TikTok W.E. Report

Weekly punch-report billing for the TikTok site (California OT/DT, pay rates + markup).

- `index.html` — the tool
- `netlify/functions/state.mjs` — the shared SAVE (Netlify Blobs storage, one backup per day)

Deploy: connect this repo to Netlify (Add new project → Import an existing project → GitHub).
No build settings to change — `netlify.toml` handles it. Every push redeploys; saved data is kept.
