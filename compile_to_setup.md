# Building Report-Generator into a Windows Setup.exe

This guide covers three steps:

1. Bundle the Python app into a standalone `.exe` with PyInstaller.
2. Wrap that into a `Setup.exe` installer with Inno Setup.
3. Publish releases on GitHub so the app can check for updates.

Do all of this on a **Windows** machine (PyInstaller builds Windows exes only when run on Windows).

---

## 1. Bundle with PyInstaller

### 1.1 Install build tools

In your project's virtual environment:

```powershell
pip install -r requirements.txt
pip install pyinstaller
```

### 1.2 Data files this project needs bundled

`run.py` is the entry point. The app also relies on these non-Python files at runtime, which PyInstaller will **not** pick up automatically:

- `rapport_template.docx`
- `rapport_mapping.json`
- `mappings.xlsx`
- `app/app_data.db`

If you add/rename data files later, update the `--add-data` list below to match.

### 1.3 Build command

From the project root:

```powershell
pyinstaller --noconfirm --onedir --windowed --name "ReportGenerator" ^
  --icon=app.ico ^
  --add-data "rapport_template.docx;." ^
  --add-data "rapport_mapping.json;." ^
  --add-data "mappings.xlsx;." ^
  --add-data "app\app_data.db;app" ^
  run.py
```

Notes:
- `--onedir` (not `--onefile`) — starts faster and is what you want anyway, since it's being installed into a folder.
- `--windowed` suppresses the console window (this is a Tkinter GUI app).
- `--icon=app.ico` is optional; remove it if you don't have an `.ico` file yet (you can convert a PNG to `.ico` with any online converter).
- Windows uses `;` as the separator in `--add-data` (`source;dest`), not `:`.
- If you get `ModuleNotFoundError` for tkinter at runtime, add `--hidden-import tkinter --hidden-import tkinter.ttk`.

### 1.4 Locate your bundled app code at runtime

PyInstaller-bundled data files end up next to the exe at runtime, not at the original relative path. If your code currently does something like:

```python
TEMPLATE_PATH = Path(__file__).resolve().parent / "rapport_template.docx"
```

this will break once frozen. Use a helper that checks for the frozen state:

```python
import sys
from pathlib import Path

def resource_path(relative_path: str) -> Path:
    if getattr(sys, "frozen", False):
        base_path = Path(sys.executable).parent
    else:
        base_path = Path(__file__).resolve().parent
    return base_path / relative_path
```

Replace any hardcoded paths to `rapport_template.docx`, `rapport_mapping.json`, `mappings.xlsx`, and `app/app_data.db` with calls to this helper. Search the codebase for these filenames before building (`grep -rn "rapport_template.docx\|rapport_mapping.json\|mappings.xlsx\|app_data.db" app`) and fix each one.

### 1.5 Build and test

```powershell
pyinstaller ReportGenerator.spec
```

(or re-run the full command above; PyInstaller generates a `.spec` file the first time you can reuse afterward — edit it directly for repeat builds instead of retyping flags).

Output lands in `dist\ReportGenerator\`. **Test this folder by copying it to a clean Windows machine or VM without Python installed** — that's the only real test that everything needed is actually bundled.

---

## 2. Wrap into Setup.exe with Inno Setup

### 2.1 Install Inno Setup

Download and install from https://jrsoftware.org/isinfo.php (free).

### 2.2 Write the installer script

Save as `installer.iss` in your project root:

```ini
[Setup]
AppId={{8B6E2F1A-9C3D-4E5A-BC11-REPLACE-WITH-NEW-GUID}}
AppName=Report Generator
AppVersion=1.0.0
AppPublisher=Your Name or Company
DefaultDirName={autopf}\ReportGenerator
DefaultGroupName=Report Generator
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=ReportGenerator-Setup-1.0.0
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\ReportGenerator.exe

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "dist\ReportGenerator\*"; DestDir: "{app}"; Flags: recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Report Generator"; Filename: "{app}\ReportGenerator.exe"
Name: "{group}\Uninstall Report Generator"; Filename: "{uninstallexe}"
Name: "{commondesktop}\Report Generator"; Filename: "{app}\ReportGenerator.exe"; Tasks: desktopicon

[Tasks]
Name: "desktopicon"; Description: "Create a desktop icon"; GroupDescription: "Additional icons:"

[Run]
Filename: "{app}\ReportGenerator.exe"; Description: "Launch Report Generator"; Flags: postinstall nowait skipifsilent
```

Important: generate your own GUID for `AppId` once (Inno Setup's Tools menu has a GUID generator, or use any online GUID generator) and **never change it again** — Windows uses it to recognize upgrades vs. fresh installs.

### 2.3 Compile the installer

Open `installer.iss` in the Inno Setup Compiler (or run from command line):

```powershell
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
```

This produces `installer_output\ReportGenerator-Setup-1.0.0.exe` — a single file that installs the app into Program Files, creates Start Menu/desktop shortcuts, and registers a proper uninstaller in "Add or Remove Programs".

### 2.4 Bump the version for each release

Each time you ship an update, change two lines in `installer.iss`:

```ini
AppVersion=1.1.0
OutputBaseFilename=ReportGenerator-Setup-1.1.0
```

and keep this in sync with `__version__` in your Python code (see step 3 below).

---

## 3. GitHub Releases (versioning)

No custom server needed — GitHub Releases acts as your version registry and file host.

### 3.1 Tag and publish a release

1. Push your code to GitHub.
2. On GitHub, go to **Releases → Draft a new release**.
3. Tag it `v1.0.0` (must match the version in your installer/app — semantic versioning: `MAJOR.MINOR.PATCH`).
4. Attach `ReportGenerator-Setup-1.0.0.exe` as a release asset (drag and drop it in).
5. Publish.

For each future update: bump the version everywhere (code, `installer.iss`), rebuild, tag a new release (`v1.1.0`), attach the new Setup.exe.

### 3.2 The endpoint your app will poll

GitHub exposes the latest release as JSON at:

```
https://api.github.com/repos/<owner>/<repo>/releases/latest
```

Relevant fields in the response:
- `tag_name` — e.g. `"v1.1.0"`
- `assets[].name` / `assets[].browser_download_url` — the Setup.exe and its direct download link

This is what the auto-update check (covered separately) will call to compare against the version running locally.

### 3.3 Single source of truth for the version number

Add to `app/__init__.py`:

```python
__version__ = "1.0.0"
```

Import this both in your update-check code and reference it when bumping `installer.iss`, so the three places a version number lives (code, installer, GitHub tag) never drift apart. Bump all three together on every release.

---

## Build checklist for each release

1. Bump `__version__` in `app/__init__.py`.
2. Run the PyInstaller command (or `pyinstaller ReportGenerator.spec`).
3. Test `dist\ReportGenerator\` on a clean machine/VM.
4. Bump `AppVersion` and `OutputBaseFilename` in `installer.iss`.
5. Compile with ISCC to get `ReportGenerator-Setup-X.Y.Z.exe`.
6. Tag a new GitHub release `vX.Y.Z` and attach the Setup.exe.