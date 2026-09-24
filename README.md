# my-python-repo

A Python project scaffold with:

- **Codespaces-ready**: opening this repo in a Codespace auto-installs everything in `requirements.txt` (see `.devcontainer/devcontainer.json`).
- **Runs on every commit**: `.github/workflows/run.yml` runs `main.py` automatically on every push, using GitHub Actions.
- **Artifact output**: anything `main.py` writes into `output/` is uploaded as a downloadable artifact on each run.

## Setup

1. Create a new repo on GitHub and push these files:
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/<your-username>/<your-repo>.git
   git push -u origin main
   ```
2. Open it in a Codespace (green "Code" button → Codespaces tab → "Create codespace"). Dependencies install automatically.
3. Every time you commit and push, the workflow runs `main.py` in the cloud automatically.

## Getting your output files

After a push:
1. Go to the **Actions** tab on GitHub.
2. Click the latest workflow run.
3. Scroll to **Artifacts** at the bottom of the run summary — download the zip named `outputs-<commit-sha>`.

## Customizing

- Add real dependencies to `requirements.txt`.
- Put your actual logic in `main.py` (or import your own modules).
- Make sure any file you want downloadable gets written under `output/`.
