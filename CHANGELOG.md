# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project intends to follow [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Unreleased

### Added

- Stable v4 API version registry and endpoint documentation.
- Winslow Technology Group branding in generated Word reports.
- Live Prism Central and Nutanix Files lifecycle data collection.

### Changed

- Executive status and alert-severity presentation is standardized.
- Support lifecycle and guest-OS compatibility data is retrieved from the Nutanix Support Portal, with local fallbacks.

## 1.0.0 - 2026-07-16

### Added

- Project documentation and installation instructions.
- Python dependency manifest.
- Automatic discovery of support CSV files in the `data/` directory.
- Timestamped execution logs, raw JSON captures, and Word reports.
- Preflight validation for required support files and output directories.
- Cluster, host, VM, CVM, alert, protection, CPU, memory, network, storage, licensing, security, lifecycle, and NCC reporting.
- Cluster Management v4.2 CVM inventory, including CVM memory and vCPU allocation.
- Physical NIC, bond, VLAN, OVS bridge, and IP assignment reporting.
- Prism Element protection domains and remote sites.
- Prism Central protection policies, replication schedules, and recovery plans.
- Security Summary and Software Lifecycle Summary with Assessment Summary navigation.
- Current LCM software and firmware inventory reporting without triggering inventory or recommendation tasks.
- `--version` command-line option.
- MIT License.
- GitHub Actions validation for Python 3.10 through 3.13.

### Changed

- The Node.js `docx` report dependency is pinned to version 9.7.1.
- CPU and memory allocation calculations include Controller VM resources.
- Report status and alert severity formatting is standardized.
- Storage, licensing, security, software lifecycle, and data protection reporting is consolidated and expanded.

### Fixed

- CVM name, memory, vCPU, host, IP address, and power-state collection.
- Physical NIC discovery and bond membership reporting across different cluster models.
- Prism Central cluster stats time-range handling.
- Active-alert collection and cluster filtering.
- Storage container reserved capacity, advertised capacity, and compression-delay reporting.
- Recovery plan discovery by using the Prism Central v3 recovery-plan API when the v4 endpoint is empty.
