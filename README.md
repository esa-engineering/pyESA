# pyESA (pyRevit Extension)

Official ESA Engineering repository for the **ESAextensions** pyRevit extension.

## Purpose

- Develop and maintain ESA pyRevit tools.
- Distribute updates to all users via **pyRevit → Update**.

---

## Repository structure
```
pyESA/                        ← repo root (= ESAextensions.extension on disk)
├── pyESA.tab/                ← main tab with all panels and tools
│   ├── Coordination.panel/
│   ├── Import-Export.panel/
│   ├── MEP.panel/
│   ├── Utilities.panel/
│   ├── Views-Sheets.panel/
│   └── WIP.panel/
├── extension.json            ← pyRevit extension metadata
├── .gitignore
└── README.md
```

> The repo root IS the extension. pyRevit clones it into a folder named
> `ESAextensions.extension` on the user's machine.

---

## Team workflow (IMPORTANT)

### Rules — mandatory for all contributors

**Rule 1 — Never commit/push directly to `main`**
`main` must always stay stable and working for all users.

**Rule 2 — All changes go through Pull Requests**
Every change must be developed on a branch and merged via PR approved by the admin.

**Rule 3 — Branch naming**
- `newfeature/<short-description>`
- `fix/<short-description>`
- `docs/<short-description>`

Examples: `newfeature/autolegend-grid`, `fix/export-crash`, `docs/install-guide`

**Rule 4 — Commit messages**
Short, plain English descriptions. No prefixes.
Examples: `Add autolegend grid tool`, `Fix export crash on RVT25`, `Update install guide`

---

## Developer setup (first time)

### 1) Clone the repo into pyRevit extensions folder
```powershell
git clone https://github.com/esa-engineering/pyESA.git "$env:APPDATA\pyRevit\Extensions\ESAextensions.extension"
```

### 2) Register the extensions folder with pyRevit
```powershell
pyrevit extensions paths add "$env:APPDATA\pyRevit\Extensions"
```

### 3) Set your Git identity for this repo
```powershell
cd "$env:APPDATA\pyRevit\Extensions\ESAextensions.extension"
git config user.name "Your Name"
git config user.email "yourname@esa-engineering.com"
```

### 4) Configure VSCode autocomplete (Revit API stubs)

Create a `pyrightconfig.json` in the repo root:
```json
{
    "pythonVersion": "3.9",
    "pythonPlatform": "Windows",
    "extraPaths": [
        "C:\\Users\\<youruser>\\AppData\\Roaming\\RevitAPI stubs\\RVT 25"
    ],
    "typeCheckingMode": "off"
}
```

> `pyrightconfig.json` is in `.gitignore` — each developer creates their own locally.

### 5) Open in VSCode
```powershell
code "$env:APPDATA\pyRevit\Extensions\ESAextensions.extension"
```

---

## Development workflow (daily use)

### Create a new branch and work on it
```powershell
git checkout main
git pull
git checkout -b newfeature/<short-description>
```

Make your changes, test in Revit (pyRevit → Reload Scripts).

### Commit and push the branch
```powershell
git add .
git commit -m "Description of what you did"
git push -u origin newfeature/<short-description>
```

### Open a Pull Request

Go to https://github.com/esa-engineering/pyESA and open a PR from your branch to `main`.
Wait for admin approval before merging.

### After merge — clean up
```powershell
git checkout main
git pull
git branch -d newfeature/<short-description>
```

---

## End user installation

> Users do NOT need a GitHub account. The extension is public.

Run this once in PowerShell:
```powershell
pyrevit extend ui ESAextensions https://github.com/esa-engineering/pyESA.git --dest="$env:APPDATA\pyRevit\Extensions" --branch=main
```

Then reload Revit. To get future updates: **pyRevit tab → Update**.