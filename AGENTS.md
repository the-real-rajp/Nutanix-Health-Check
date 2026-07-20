# AGENTS.md

## Project scope

This repository contains the Nutanix Health Check report generator. The Python
source distribution and the self-contained Windows x64 package are two builds
of the same application and must remain functionally consistent.

## Source of truth

- `nutanix_health_check.py` is the primary application source.
- `main` contains the latest stable, tested baseline.
- `v1.0.0` is the first stable release and must not be rewritten.
- New v1.1 development work belongs on `develop/v1.1.0` or a short-lived
  feature branch created from it.
- Do not modify, move, or delete release tags without explicit user approval.

## Change workflow

- Make one focused enhancement at a time.
- Do not commit, push, merge, tag, or publish a release until the user confirms
  that the change was tested successfully.
- Preserve all unrelated working-tree changes.
- Before handing a script to the user, run syntax and CLI validation.
- Test report changes against all three representative clusters when sample
  data is available:
  - `NX-8170-01`
  - `NX-8035-01`
  - `NX-8035-02`
- After approval, use a pull request and squash merge into the protected target
  branch.
- Commits must use the configured `the-real-rajp` GitHub identity.

## API safety

- Health-check collection must remain read-only.
- Do not invoke LCM inventory, recommendation, upgrade, remediation, or other
  action endpoints.
- Do not acknowledge, resolve, modify, or delete Prism alerts or configuration.
- POST requests are allowed only for read-only list/search APIs that require a
  POST body, such as Prism Central v3 list endpoints.
- Prefer current Prism Central v4 APIs. Retain v3 or Prism Element fallbacks
  where they are required for confirmed compatibility.
- Document new API families, versions, methods, and endpoint patterns in
  `docs/API_VERSION_MATRIX.md`.

## Report compatibility

- Preserve existing working report sections unless the requested change
  explicitly modifies them.
- Keep report terminology, capitalization, status colors, and table layout
  consistent across sections.
- Do not silently replace unavailable data with inferred values.
- Use `N/A` or omit an optional table when the API does not provide reliable
  data, according to the established behavior of that report section.
- Keep timestamped DOCX, JSON, and log filenames Windows- and macOS-safe.

## Validation

Run the relevant checks before asking the user to test:

```bash
python -m py_compile nutanix_health_check.py
python nutanix_health_check.py --version
python nutanix_health_check.py --help
```

For Windows packaging changes, also validate:

```powershell
.\build-windows.ps1
```

The extracted Windows package must run without separately installed Python,
Node.js, npm packages, or CSV support files.

## Files and sensitive data

- Do not commit generated DOCX reports, raw JSON captures, logs, temporary
  report-builder files, virtual environments, Node modules, bundled runtimes,
  or PyInstaller build output.
- Treat uploaded reports and raw JSON as customer-sensitive infrastructure
  data even when the repository is public.
- Never store Prism credentials, passwords, tokens, or private certificates in
  source, documentation, logs, examples, commits, or release assets.
- The required public support data belongs under `data/`:
  - `OS_Compatibility_Matrix.csv`
  - `NOS_EOL_information_list.csv`

## Windows distribution

- Keep `nutanix_health_check.py` available as the editable Python source.
- Keep Windows packaging source files in Git, including the PyInstaller spec,
  build script, and launcher.
- Keep generated `vendor/`, `build/`, and `dist/` directories out of Git.
- Publish the tested Windows ZIP as a GitHub release asset instead of committing
  the binary to the repository.
