# Windows build and packaging

This foundation uses Python's standard library. It can be run directly with Python 3.11+.

## Run the desktop interface

1. Install Python 3.11+ on a Windows test computer and include Tcl/Tk when installing.
2. Copy this folder locally.
3. Run `setup-first-run.bat`, then confirm the local database path and source identity.
4. Run `python bridge_ui.py` to open the desktop interface.

## Build a standalone executable later

After testing on Windows, optionally install PyInstaller on the build computer and run:

```bat
python -m pip install pyinstaller
pyinstaller --noconfirm --clean --onefile --windowed --name EasyBuysAroniumBridge bridge_ui.py
```

The executable is generated in `dist\EasyBuysAroniumBridge.exe`. Keep `bridge.py`, `config.example.json`, and the deployment instructions beside the executable, or package them deliberately as application data. The standalone executable has not been built or tested by this repository change.

Do not distribute a production installer until the bridge has been tested with real Aronium schema copies and Windows security/backup behaviour has been reviewed.
