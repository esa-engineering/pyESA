# pyESA (pyRevit Extension)

Official ESA Engineering repository for the **ESAextensions.extension** pyRevit extension.

## Purpose
- Develop and maintain ESA pyRevit tools.
- Distribute updates to users via **pyRevit → Update** (git pull on local clones).

## Repository structure
- `ESAextensions.extension/` — the actual pyRevit extension (tabs/panels/buttons/scripts)
- `.github/` — PR template and contribution guidelines

> pyRevit loads the `.extension` folder. `.tab` folders live inside it.

---

## Team workflow (IMPORTANT)
This repository is **private** and branch protections may not be technically enforced depending on the organization plan.
For this reason, the following process is **mandatory** for all contributors.

### Rule 1 — Never commit/push directly to `main`
`main` must always stay stable and releasable.

### Rule 2 — All changes go through Pull Requests
Every change must be done on a branch and merged via PR.

### Rule 3 — Branch naming
Use only the following patterns:
- `newfeature/<short-description>`
- `fix/<short-description>`
- `docs/<short-description>`

Examples:
- `newfeature/legend-grid-generator`
- `fix/update-timeout`
- `docs/installation-notes`

### Rule 4 — Commit messages
Use **short, plain descriptions** (no prefixes).
Examples:
- `Add legend grid generator`
- `Fix update command crash`
- `Update installation notes`

---

## Development workflow (Git)
### 1) Create a branch
```bash
git checkout main
git pull
git checkout -b newfeature/<name>
