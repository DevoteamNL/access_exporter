# 01 – Application icon

Status: **done** · Size: S · Live data: no

## Goal

Give Access Exporter an icon in the same style as iAsset Syncer (`iAsset_Syncer/iasset/assets/iasset-syncer.svg`):
rounded square with blue→purple gradient, lavender top bracket and mint bottom bracket, with the orange place marker
replaced by an orange **database** (stacked cylinder). Use it for the window, taskbar and the `.exe`.

## Draft SVG (`assets/access-exporter.svg`)

Background and brackets copied 1:1 from iAsset Syncer; only the centre glyph changes. Proportions to be tuned
visually at 16, 32 and 256 px.

```svg
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 128 128">
  <defs>
    <linearGradient id="background" x2="1" y2="1">
      <stop stop-color="#1855c9"/>
      <stop offset="1" stop-color="#482b9e"/>
    </linearGradient>
  </defs>
  <rect x="4" y="4" width="120" height="120" rx="28" fill="url(#background)"/>
  <!-- top (#cbb5ff) and bottom (#78e3b7) bracket paths: copy unchanged from iasset-syncer.svg -->
  <g transform="translate(64 64) scale(.9) translate(-64 -64)" fill="#ff9854">
    <ellipse cx="64" cy="50" rx="17" ry="6"/>
    <path d="M47 54 A17 6 0 0 0 81 54 V64 A17 6 0 0 1 47 64 Z"/>
    <path d="M47 68 A17 6 0 0 0 81 68 V78 A17 6 0 0 1 47 78 Z"/>
  </g>
</svg>
```

## Steps

1. Add `assets/access-exporter.svg` (above).
2. Add `tools/generate_app_icon.py`, adapted from iAsset Syncer's generator (QtSvg, offscreen, sizes 16–256 →
   `assets/access-exporter.ico` + 1024 px `assets/access-exporter.png`). Commit the generated files.
3. `access_exporter_pyqt6.py`: add `resource_path(name)` (handles PyInstaller `sys._MEIPASS`), call
   `app.setWindowIcon(QIcon(resource_path("assets/access-exporter.ico")))` in `__main__`, and set the Windows
   AppUserModelID (`SetCurrentProcessExplicitAppUserModelID`) so the taskbar shows the icon.
4. PyInstaller: add `--icon assets/access-exporter.ico --add-data "assets;assets"` to README (EN + NL) and
   `.github/workflows/release-assets.yml`.

## Tests

- `tests/test_icon.py`: SVG is valid (`QSvgRenderer.isValid()`), `.ico` exists and contains the expected sizes
  (parse the ICO header), `resource_path` resolves from source and with a fake `sys._MEIPASS`.

## Risks

- `PyQt6.QtSvg` is only needed by the generator (dev-time) → no new runtime dependency.
- Missing `--add-data` gives an exe without window icon; covered by the workflow change.

## Open question

- Database colour: same orange `#ff9854` as the pin (recommended), or different?
