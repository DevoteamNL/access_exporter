# 06 – README completeness (EN + NL)

Status: **approved** · Size: S · Live data: no

## Gaps in the current README.md

1. Repo/release links use `cramor-devo/access_exporter`; the repository is `DevoteamNL/access_exporter` (confirm).
2. UI supports French, README has EN + NL only.
3. "Dark theme" only – changes with plan 02.
4. Missing sections: **Screenshot**, **Troubleshooting** (driver missing, 32/64-bit mismatch, file locked by Excel,
   database open in Access), **Where settings are stored** (`HKCU\Software\AccessExporter`), **Limitations**
   (attachment/OLE/multi-valued fields, saved queries not exported), **Changelog** (no `CHANGELOG.md`).
5. *Development* documents `unittest`; the repo's Checks script uses `pytest` – document both or align.
6. Build command lacks icon/asset flags (01) and `openpyxl` (04).
7. EN and NL sections must stay mirrored on every change.

## Steps

1. Quick-fix commit: items 1, 4 (troubleshooting, settings, limitations), 5.
2. Each feature PR (01–05) updates EN and NL with the exact UI labels.
3. Add `CHANGELOG.md` with 1.1.0 and 1.2.0.
4. Final pass: screenshots in `docs/images/` (dark + light).

## PO decisions

- Three files: `README.md` (EN), `README.nl.md`, `README.fr.md`, kept in sync, each linking the others.
- Links point to `DevoteamNL/access_exporter`.
- Rewrite for legibility; start with a plain-language introduction for non-developers (consultants): what it does,
  when to use it, how to get started with the `.exe` in 3 steps. Developer content moves to a section at the end.
