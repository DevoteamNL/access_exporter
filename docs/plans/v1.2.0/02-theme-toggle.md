# 02 – Light/dark theme toggle

Status: **done** · Size: M · Live data: no

## Goal

Let the user switch between dark and light theme (optionally "System") from the top bar next to the language
selector; the choice is saved and applied live to the main window and all dialogs.

## Findings

- `AccessExporterGUI.apply_stylesheet()` holds one hard-coded dark QSS.
- Colours outside the QSS: `KLEUR_INFO` (path labels via `setStyleSheet`), `zet_tabellen_status` colours,
  `STATUS_KLEUREN` and fallback `#E0E0E0` in `update_table_row`. These must come from the active palette.
- Fusion style is already set → `QPalette` + QSS works for both themes.
- System detection: `QGuiApplication.styleHints().colorScheme()` (Qt ≥ 6.5) → `PyQt6>=6.5` in `requirements.txt`
  if "System" is chosen.

## Steps

1. New `themes.py` (no widgets): `THEMES = {"dark": {...}, "light": {...}}` with named colours (background, surface,
   text, border, accent, ok, error, info, busy) and `build_stylesheet(theme) -> str`.
2. Replace `apply_stylesheet()` with `apply_theme(name)`: sets the `QApplication` stylesheet (dialogs inherit it) and
   re-colours path labels, the tables-status label and queue status cells.
3. Add combo **Theme** with **Dark**, **Light** (and **System** if approved) next to **Language**; saved in
   `QSettings` key `thema`, loaded in `laad_instellingen` (unknown value → default).
4. Translations NL/EN/FR in `translations.py` (`Thema`, `Donker`, `Licht`, `Systeem`).

## Tests

- `tests/test_theme.py` (offscreen): switching changes the application stylesheet; status cells and labels use the
  theme colours; preference persists via a temporary `QSettings` file; invalid stored value falls back.
- `test_languages.py` stays green; catalog test covers the new strings.

## Risks

- Contrast of status colours on white → darker variants in the light palette.

## PO decision

- **System / Light / Dark**, default **System** (follows Windows, live on `colorSchemeChanged`); `PyQt6>=6.5`.
