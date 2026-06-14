# Chromatic Guide

Personal site — **chromaticguide.com**

Balancing people, work, and AI. Light static HTML, hosted on GitHub Pages.

## Structure

```
chromaticguide/
  index.html          Home
  about.html          About
  resources/          Guides, ideas, tool notes (stubs for now)
  assets/css/site.css Shared styles
  assets/favicon.svg
  CNAME               chromaticguide.com
```

## Deploy to GitHub

From this folder (after cloning your repo locally):

```powershell
cd path\to\chromaticguide
git init
git remote add origin https://github.com/melonie1616-ai/chromaticguide.git
git add .
git commit -m "Add light theme site: home, about, resources"
git branch -M main
git pull origin main --rebase
git push -u origin main
```

If the repo already has commits (e.g. README), pull first, copy files in, then add/commit/push.

Or copy files into your existing clone and push.

## Local preview

Open `index.html` in a browser, or:

```powershell
python -m http.server 8080
```

Then visit http://localhost:8080
