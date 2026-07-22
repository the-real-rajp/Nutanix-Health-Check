# Nutanix Health Check

`nutanix_health_check.py` connects to Nutanix Prism Central, collects cluster inventory and health data through REST APIs, and generates a Microsoft Word health-check report for each registered cluster.

Current release: **v1.0.0**

## Report coverage

The report includes:

- Cluster and host inventory
- Controller VM and user VM inventory
- Active alerts and recommended actions
- CPU and memory utilization and allocation
- Network, bond, VLAN, and physical NIC information
- Storage capacity, container configuration, and encryption
- Licensing, protection domains, protection policies, and recovery plans
- Security configuration and active security alerts
- AOS software lifecycle and current LCM firmware inventory

The complete API-family and endpoint inventory is documented in
[`docs/API_VERSION_MATRIX.md`](docs/API_VERSION_MATRIX.md).

## Requirements

### Python source distribution

- Python 3.10 or later
- Node.js 18 or later
- Network access to Prism Central on port `9440`
- A Prism Central account with permission to read the required inventory and statistics
- The support files in [`data/`](data/):
  - `OS_Compatibility_Matrix.csv`
  - `NOS_EOL_information_list.csv`
- The report logo at `images/winslow-technology-group-logo.png`

The script installs the pinned Node.js `docx` package (`9.7.1`) locally on its
first report run if that exact version is not already available. Pinning the
package prevents an upstream release from unexpectedly changing report
rendering. Matplotlib is used to generate the CPU, memory, and storage charts.

### Windows portable distribution

The Windows x64 release is self-contained and does not require Python,
Node.js, npm, or the support CSV files to be installed separately. Download
`Nutanix-Health-Check-1.0.0-Windows-x64.zip` from the GitHub release, extract
the complete archive, and run `Run-Nutanix-Health-Check.cmd`. Reports, raw JSON
captures, and logs are written beneath the extracted `output` directory.

Keep `NutanixHealthCheck.exe`, `Run-Nutanix-Health-Check.cmd`, and the
`_internal` directory together. The executable depends on the bundled runtime
under `_internal`.

## Installation

### Python source

Clone the repository and create a virtual environment:

```bash
git clone https://github.com/the-real-rajp/Nutanix-Health-Check.git
cd Nutanix-Health-Check

python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### Building the Windows portable package

On a Windows x64 build computer with Python, Node.js, npm, and PyInstaller
available, activate the Python build environment and run:

```powershell
.\build-windows.ps1
```

The build script validates the application, bundles Node.js and the pinned
`docx` package, runs PyInstaller, and creates:

```text
dist\Nutanix-Health-Check-1.0.0-Windows-x64.zip
```

Build products and downloaded runtimes under `build/`, `dist/`, and `vendor/`
are intentionally excluded from Git. The tested ZIP is distributed as a
GitHub release asset.

## Usage

### Recommended: interactive mode

Run the script without arguments:

```bash
python nutanix_health_check.py
```

The script starts with **Preflight Validation**. Before requesting Prism Central credentials, it confirms that these required support files are available:

- `data/OS_Compatibility_Matrix.csv`
- `data/NOS_EOL_information_list.csv`
- `images/winslow-technology-group-logo.png`

A successful preflight looks similar to:

```text
------------------------------------------------------------
Nutanix Health Check - Preflight Validation
------------------------------------------------------------

Checking required support files...

  [OK] OS_Compatibility_Matrix.csv
  [OK] NOS_EOL_information_list.csv
  [OK] winslow-technology-group-logo.png

All required support files found.
Proceeding to Prism Central connection...
```

After validation, the script interactively prompts for:

1. Prism Central IP address or FQDN
2. API port, with `9440` as the default
3. Prism Central username
4. Password, entered through a hidden prompt
5. Customer or organization name for the report

The script then tests the Prism Central connection, discovers registered clusters, collects each cluster's data, and generates the raw JSON and Word health-check report.

If a required CSV or the report logo is missing, preflight stops before any connection information is requested and lists the expected filename.

### Optional command-line mode

Prism Central, user, customer, and output settings can be supplied as arguments while still allowing the script to prompt securely for the password:

```bash
python nutanix_health_check.py \
  --host pc.example.com \
  --user admin \
  --customer "Customer Name" \
  --output-dir reports
```

Avoid supplying `--password` directly when possible because command-line values may be retained in shell history or visible to other processes.

### Generate from saved JSON

Generate a report from an existing raw JSON capture without connecting to Prism Central:

```bash
python nutanix_health_check.py \
  --from-json reports/CLUSTER_raw.json \
  --customer "Customer Name" \
  --output-dir reports
```

The script automatically searches the project root and `data/` for both required CSV files. The report logo is loaded from `images/` beside the script or from the bundled Windows resources. Custom CSV paths can be supplied with:

```bash
--os-compat-csv /path/to/OS_Compatibility_Matrix.csv
--aos-eol-csv /path/to/NOS_EOL_information_list.csv
```

## Generated files

Depending on the selected mode, the script creates:

- `<cluster>_raw_YYYY-MM-DD_HH-MM-SS.json`
- `<cluster>_Health_Check_YYYY-MM-DD_HH-MM-SS.docx`
- `logs/Nutanix_Health_Check_YYYY-MM-DD_HH-MM-SS.log`

The preflight creates and validates the `logs` folder inside the selected
output directory. The timestamped execution log captures preflight results,
collection progress, warnings, generated file paths, errors, and completion
time. Its filename avoids characters that are invalid on Windows and is also
safe to use on macOS.

Generated reports, raw captures, temporary report-builder files, Python caches, and local Node.js packages are excluded by `.gitignore`.

## Version and validation

Display the installed script version:

```bash
python nutanix_health_check.py --version
```

GitHub Actions validates Python 3.10 through 3.13 by installing the declared
dependencies, compiling the script, and checking the command-line interface.

## Data and security considerations

- Raw JSON captures and generated reports can contain infrastructure names, IP addresses, alerts, and other sensitive configuration data. Store and share them appropriately.
- Review the licensing and redistribution terms for the CSV support files before publishing them in a public repository.
- Validate recommendations against current Nutanix documentation and your organization's operational requirements before implementing changes.

## License

This project is available under the [MIT License](LICENSE).

## Project status

Version 1.0.0 is the first stable reporting baseline. Continue testing future
changes against representative clusters before promoting a new release.
