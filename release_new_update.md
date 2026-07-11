# Releasing a New Update

## 1. Bump the version

**`app/__init__.py`**
```python
__version__ = "1.0.2"  # increment this
```

**`installer.iss`** — update two lines:
```ini
AppVersion=1.0.2
OutputBaseFilename=ReportGenerator-Setup-1.0.2
```

---

## 2. Rebuild the exe and installer

```powershell
pyinstaller run.spec
& "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
```

Output: `installer_output\ReportGenerator-Setup-1.0.2.exe`

---

## 3. Push to GitHub

```powershell
git add .
git commit -m "Release v1.0.2"
git tag v1.0.2
git push origin main
git push origin v1.0.2
```

---

## 4. Publish the GitHub Release

1. Go to **https://github.com/alexbalak21/Report-Generator/releases/new**
2. Under **Choose a tag** → select `v1.0.2`
3. Title: `v1.0.2`
4. Drag and drop `installer_output\ReportGenerator-Setup-1.0.2.exe` into the assets area
5. Click **Publish release**

---

## Done

Users still on an older version will be prompted to update the next time they launch the app.