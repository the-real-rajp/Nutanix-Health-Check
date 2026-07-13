# Nutanix Health Check

`nutanix_health_check.py` connects to Nutanix Prism Central, collects cluster inventory and health data through REST APIs, and generates a Microsoft Word health-check report for each registered cluster.

## Report coverage

The report includes:

- Cluster and host inventory
- Controller VM and user VM inventory
- Active alerts and NCC findings
- CPU and memory utilization and allocation
- Network, bond, VLAN, and physical NIC information
- Storage capacity, container configuration, and encryption
- Licensing and protection-domain information
- Security configuration and active security alerts
- AOS software lifecycle information

## Requirements

- Python 3.10 or later
- Node.js 18 or later
- Network access to Prism Central on port `9440`
- A Prism Central account with permission to read the required inventory and statistics
- The support files in [`data/`](data/):
  - `OS_Compatibility_Matrix.csv`
  - `NOS_EOL_information_list.csv`

The script installs the Node.js `docx` package locally on its first report run if the package is not already available. Matplotlib is used to generate the CPU, memory, and storage charts.

## Installation

Clone the repository and create a virtual environment:

```bash
git clone https://github.com/the-real-rajp/Nutanix-Health-Check.git
cd Nutanix-Health-Check

python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

## Usage

Interactive mode:

```bash
python nutanix_health_check.py
```

Specify Prism Central and the user name while allowing the script to prompt securely for the password:

```bash
python nutanix_health_check.py \
  --host pc.example.com \
  --user admin \
  --customer "Customer Name" \
  --output-dir reports
```

Avoid supplying `--password` directly when possible because command-line values may be retained in shell history or visible to other processes.

Generate a report from an existing raw JSON capture without connecting to Prism Central:

```bash
python nutanix_health_check.py \
  --from-json reports/CLUSTER_raw.json \
  --customer "Customer Name" \
  --output-dir reports
```

The script automatically searches the project root and `data/` for both required CSV files. Custom paths can be supplied with:

```bash
--os-compat-csv /path/to/OS_Compatibility_Matrix.csv
--aos-eol-csv /path/to/NOS_EOL_information_list.csv
```

## Generated files

Depending on the selected mode, the script creates:

- `<cluster>_raw.json`
- `<cluster>_Health_Check.docx`

Generated reports, raw captures, temporary report-builder files, Python caches, and local Node.js packages are excluded by `.gitignore`.

## Data and security considerations

- Raw JSON captures and generated reports can contain infrastructure names, IP addresses, alerts, and other sensitive configuration data. Store and share them appropriately.
- Review the licensing and redistribution terms for the CSV support files before publishing them in a public repository.
- Validate recommendations against current Nutanix documentation and your organization's operational requirements before implementing changes.

## Project status

The project is under active development. Test changes against representative clusters before using a new revision as the reporting baseline.

