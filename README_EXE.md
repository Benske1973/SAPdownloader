## EXE maken (Windows)

Belangrijk: een **Windows `.exe`** bouw je het best op **Windows** (PyInstaller is OS-specifiek).

### Vereisten
- Python 3.10+ (aanbevolen)
- Microsoft Excel (voor de COM-automatisatie)
- Microsoft Edge geïnstalleerd (de app gebruikt Playwright met `channel="msedge"`)

### Build stappen
Open **Command Prompt** in deze map en run:

```bat
build_exe.bat
```

Na afloop staat de EXE hier:
- `dist\EquansSAPDownloader.exe`

### Troubleshooting
- **Playwright error over browsers**: normaal niet nodig omdat we `msedge` gebruiken, maar als je toch errors krijgt:
  - run `python -m playwright install`
- **pywin32/COM issues**:
  - zorg dat je in dezelfde Python runt waarmee je buildt
  - herinstalleer: `pip install --force-reinstall pywin32`
