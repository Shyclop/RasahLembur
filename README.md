RaSahLembur Excel Automation

Goal: produce a single-file Windows executable and an installer so end-users can run the app without installing Python or dependencies.

Prerequisites (build machine):
- Windows machine (same arch as target, e.g., x64)
- Python 3.11+ installed
- Inno Setup (for building installer) if you want an installer

Quick build steps (recommended reproducible flow):

1) Create and activate virtual environment
```powershell
cd d:\ART\MINE\Apps\Projects\erpauto
python -m venv .venv
.venv\Scripts\Activate.ps1
```

2) Install dependencies and pin them
```powershell
pip install --upgrade pip
pip install -r requirements.txt
pip freeze > requirements.txt
```

3) Build single EXE with PyInstaller
```powershell
pyinstaller --onefile --windowed --name "RaSahLemburExcel" auto.py
```
Output: `dist\RaSahLemburExcel.exe`

4) (Optional) Create an installer with Inno Setup
- Use `installer.iss` included in this repo as a starting point.
- Open it with Inno Setup Compiler and build.

5) (Recommended) Code-sign the EXE & installer
- Obtain an Authenticode certificate and use `signtool.exe`:
```powershell
signtool sign /a /tr http://timestamp.digicert.com /td sha256 /fd sha256 path\to\RaSahLemburExcel.exe
```

Notes and troubleshooting:
- Some dependencies include native extensions and require the Microsoft Visual C++ Redistributable. Include `vcredist_x64.exe` in your installer if needed.
- If PyInstaller misses files (Tkinter data files), you may need to pass `--add-data` or use a spec file. I can generate a spec file if you need it.
- If you want, I can attempt the build for you here again, or produce the installer script and spec file ready to run locally.

If you'd like, I can now:
- generate a PyInstaller spec that includes the tkinter runtime files and icons,
- create `installer.iss` (already included) with VC++ redistributable inclusion,
- or retry the automated build here and debug errors.
Which do you prefer?