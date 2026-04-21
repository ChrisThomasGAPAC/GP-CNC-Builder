# GP CNC Builder

GP CNC Builder is a desktop layout tool for planning gypsum board sample cuts for DTP/STP test procedures. It generates board layouts, validates spacing and procedure requirements, supports manual part placement, and exports DXF/SVG/CSV/PDF/job-package outputs.

## Highlights

- Multiple sample request rows for mixed boards.
- Rule-aware placement for DTP/STP procedures.
- Fill Board layout spacing for cleaner board use.
- 2 inch shop edge inset for non-edge-shear samples.
- STP318 Edge Shear formed-edge exception.
- Manual add, move, duplicate, rotate, lock, measure, highlight, and text tools.
- Saved sheets by signed-in email profile.
- DXF export for current sheets and saved sheets.
- PyInstaller spec for creating a no-admin Windows executable.

## Run From Source

```powershell
python dtp.py
```

## Build Windows EXE

```powershell
python -m PyInstaller --noconfirm --clean --onefile --windowed --name "GP CNC Builder" --icon "assets\gp_cnc_builder.ico" --add-data "assets;assets" dtp.py
```

The built executable is generated under `dist/`.

## Repository Layout

- `dtp.py` - main Tkinter application.
- `assets/` - app icons and GP logo assets.
- `examples/` - sample DXF exports.
- `GP CNC Builder.spec` - PyInstaller build configuration.

Generated folders such as `build/`, `dist/`, and `__pycache__/` are intentionally ignored.
