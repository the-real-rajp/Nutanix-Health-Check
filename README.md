# Nutanix Health Check

`nutanix_health_check.py` connects to Nutanix Prism Central, collects cluster inventory and health data through REST APIs, and generates a Microsoft Word health-check report for each registered cluster.

Development version: **v1.1.0**

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

- Python 3.10 or later
- Node.js 18 or later
- Network access to Prism Central on port `9440`
- A Prism Central account with permission to read the required inventory and statistics
- HTTPS access to the Nutanix Support Portal for current AOS, Prism Central,
  and Nutanix Files lifecycle information and guest-OS compatibility data
- Optional offline CSV fallbacks in [`data/`](data/):
  - `OS_Compatibility_Matrix.csv`
  - `NOS_EOL_information_list.csv`
- HTTPS access to the GitHub-hosted report logo; a bundled copy under `images/` is used automatically if the download is unavailable

The script installs the pinned Node.js `docx` package (`9.7.1`) locally on its
first report run if that exact version is not already available. Pinning the
package prevents an upstream release from unexpectedly changing report
rendering. Matplotlib is used to generate the CPU, memory, and storage charts.

## Installation

Clone the repository and create a virtual environment:

```bash
git clone --branch develop/v1.1.0 --single-branch https://github.com/the-real-rajp/Nutanix-Health-Check.git
cd Nutanix-Health-Check

python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Usage

### Recommended: interactive mode

Run the script without arguments:

```bash
python nutanix_health_check.py
```

The script starts with **Preflight Validation**. Before requesting Prism Central credentials, it confirms that current report-support data is available from:

- [Nutanix End of Support Life Information - Software Releases](https://portal.nutanix.com/page/documents/eol/list?type=aos)
- [Prism Central End of Support Life Information](https://portal.nutanix.com/page/documents/eol/list?type=pc)
- [Nutanix Files End of Support Life Information](https://portal.nutanix.com/page/documents/eol/list?type=files)
- [Nutanix Guest OS Compatibility Matrix](https://portal.nutanix.com/page/compatibility-interoperability-matrix/guestos/compatibility)
- Winslow Technology Group report logo from the configured GitHub URL, with `images/winslow-technology-group-logo.png` as its fallback

If either Nutanix portal API is unavailable, the script automatically uses its
packaged CSV fallback under `data/`.

A successful preflight looks similar to:

```text
------------------------------------------------------------
Nutanix Health Check - Preflight Validation
------------------------------------------------------------

Checking report support data...

  [OK] Guest OS compatibility data (Nutanix Support Portal)
  [OK] AOS lifecycle data (Nutanix Support Portal)
  [OK] Prism Central lifecycle data (Nutanix Support Portal)
  [OK] Nutanix Files lifecycle data (Nutanix Support Portal)
  [OK] Report logo downloaded from https://raw.githubusercontent.com/...

All report support data is available.
Proceeding to Prism Central connection...
```

After validation, the script interactively prompts for:

1. Prism Central IP address or FQDN
2. API port, with `9440` as the default
3. Prism Central username
4. Password, entered through a hidden prompt
5. Customer or organization name for the report

The script then tests the Prism Central connection, discovers registered clusters, collects each cluster's data, and generates the raw JSON and Word health-check report.

If AOS or guest-OS portal data and its corresponding CSV fallback are both
unavailable, or the report logo is unavailable from both the web and bundled
fallback, preflight stops before requesting connection information. Prism
Central and Files lifecycle lookups are optional and produce a warning when
their portal data cannot be reached.

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

The script retrieves AOS, Prism Central, and Nutanix Files lifecycle data and
guest-OS compatibility data from the official Nutanix Support Portal APIs. It automatically searches the project
root and `data/` for offline CSV fallbacks. During report generation it
downloads the report logo from GitHub and embeds it in the DOCX; the completed
report does not depend on the URL when opened. If the download fails, the
script uses the copy under `images/` or the bundled Windows resources. Custom
fallback CSV paths can be supplied with:

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
- Lifecycle and compatibility data is read from the official Nutanix Support Portal; packaged AOS and guest-OS CSV files are offline fallbacks and should be refreshed periodically.
- Validate recommendations against current Nutanix documentation and your organization's operational requirements before implementing changes.

## License

This project is available under the [MIT License](LICENSE).

## Project status

Version 1.0.0 is the first stable reporting baseline. Continue testing future
changes against representative clusters before promoting a new release.
