# Chromatic Guide

**https://chromaticguide.com** — balancing people, work, and AI.

## Site map

| Path | Purpose |
|------|---------|
| `index.html` | Home |
| `about.html` | About |
| `resources/` | Take-home hub (PMI talk artifacts) |
| `resources/prompt-library.html` | PM AI Prompt Library |
| `resources/what-stays-with-the-pm.html` | What Stays With the PM |
| `resources/spec-template.html` | Before You Build, Write the Spec |
| `resources/experiment-playbook.html` | AI Experiment Playbook |
| `presentation/ai-for-project-managers.html` | **Horizontal slide deck** (← → keys) — Partnering with AI: 12 slides + live Cursor demo |
| `pmi-presenter-kit.html` | **All presenter materials** — run of show, speaker notes, ATLAS script, Cursor demo, Zoom chat, Q&A stories (expand/collapse sections) |
| `pmi-zoom-chat-message.txt` | Plain-text copy of Zoom chat message (also in presenter kit) |

Legacy redirects: `resources/guardrails.html` → what-stays-with-the-pm; `resources/week-month-checklist.html` → experiment-playbook.

## Deploy

Upload the contents of this folder to [melonie1616-ai/chromaticguide](https://github.com/melonie1616-ai/chromaticguide) on the `main` branch (GitHub web upload or push as `melonie1616-ai`).

**Important:** Upload each file to its **correct path**. A common mistake is replacing the root `index.html` with the Resources page — the home page headline must say *“Balancing people, work, and AI…”*, not “Resources”.

HTML/CSS alone is not enough — pages reference asset files that must be uploaded too:

| Upload to repo path | Used on |
|---------------------|---------|
| `assets/logo-wheel.svg` | Header logo (mini Pantone wheel with white borders) |
| `assets/pantone-wheel.svg` | Home, About, Resources (circular color wheel, bottom-right) |
| `assets/favicon.svg` | Browser tab icon |

**Presentation illustrations** are embedded **inside** `presentation/ai-for-project-managers.html` (inline SVG). You do **not** need to upload `presentation/images/` for the deck to show artwork — only upload the HTML file after edits.

Optional: `presentation/images/` remains the source folder if you re-run `presentation/scripts/embed_svgs.py` after changing artwork.

## Present the deck

Open `presentation/ai-for-project-managers.html` in a browser full screen (F11). Use **← →**, **Space**, or on-screen buttons. Link from home for attendees.

## Local preview

```powershell
cd "path\to\chromaticguide"
python -m http.server 8080
```
