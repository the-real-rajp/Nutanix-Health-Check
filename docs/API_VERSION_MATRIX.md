# Nutanix API and Version Matrix

This chart documents the Nutanix REST API families used by Nutanix Health
Check. Endpoint versions are listed in the order attempted by the script.
Unless otherwise noted, Prism Central requests use HTTPS port `9440` with the
base path `/api`.

## API family summary

| Platform | API family | Versions used | Role in the health check |
|---|---|---|---|
| Prism Central | Cluster Management | v4.2, v4.1, v4.0.b1, v4.0 | Cluster, host, CVM, storage, NIC, AHV, CPU, memory, and storage statistics |
| Prism Central | Security | v4.1, v4.0 | Cluster security-hardening state |
| Prism Central | Monitoring | v4.2, v4.1, v4.0 | Active cluster alerts |
| Prism Central | VMM AHV | v4.2, v4.1, v4.0 | User VM inventory and allocation |
| Prism Central | Data Policies | v4.2, v4.1 | Protection policies and current-generation recovery plans |
| Prism Central | Data Protection | v4.2, v4.1, v4.0 | Recovery-plan jobs and older protection-policy compatibility |
| Prism Central | Networking | v4.2, v4.1, v4.0.b1, v4.0 | Subnets, virtual switches, physical NICs, bonds, uplinks, and bridges |
| Prism Central | Licensing | v4.1, v4.0 | Entitlements, compliance, and violations |
| Prism Central | Lifecycle | v4.2, v4.1, v4.0 | Read-only LCM software and firmware inventory |
| Prism Central | Operations Management | v4.0 | NCC health-check alerts |
| Prism Central | Nutanix legacy API | v3.1/v3 | Alert, VM, recovery-plan, and recovery-plan-job compatibility fallbacks |
| Prism Element | Prism REST | v2.0 | Cluster, host, CVM, alert, network, Protection Domain, and Remote Site fallbacks |
| Prism Element | Prism REST | v1 | Legacy cluster CPU statistics fallback |

## Endpoint chart

| Data collected | Method | Preferred endpoint pattern | Versions/fallbacks |
|---|---|---|---|
| Connection and cluster discovery | GET | `/clustermgmt/{version}/config/clusters` | Stable versions are attempted newest-first: v4.2, v4.1, v4.0 |
| Cluster configuration | GET | `/clustermgmt/{version}/config/clusters/{clusterExtId}` | v4.2, v4.1, v4.0.b1, v4.0 |
| Hosts and AHV version | GET | `/clustermgmt/{version}/config/clusters/{clusterExtId}/hosts` | `host-nodes`, `nodes`, and global filtered host lists are fallbacks |
| Controller VMs | GET | `/clustermgmt/{version}/config/clusters/{clusterExtId}/cvms` | Per-CVM detail uses `/cvms/{cvmExtId}`; PE v2 and PC v3 VM inventory are fallbacks |
| User VMs | GET | `/vmm/{version}/ahv/config/vms` | v4.2, v4.1, v4.0; filtered by cluster; PC v3 `/nutanix/v3/vms/list` is a CVM fallback |
| Active alerts | GET | `/monitoring/{version}/serviceability/alerts` | v4.2, v4.1, v4.0; PC v3 and PE v2 are compatibility fallbacks |
| Security hardening | GET | `/security/{version}/report/security-summaries` | v4.1, v4.0; PE v2 `/cluster/` supplements the response |
| Storage containers | GET | `/clustermgmt/{version}/config/storage-containers` | v4.2, v4.1, v4.0.b1, v4.0 |
| Container statistics | GET | `/clustermgmt/{version}/stats/storage-containers/{containerExtId}` | Prefer the stats link returned by the container object |
| Cluster statistics | GET | `/clustermgmt/{version}/stats/clusters/{clusterExtId}` | Required start/end time; v4.2, v4.1, v4.0.b1, v4.0 |
| Host statistics | GET | `/clustermgmt/{version}/stats/hosts/{hostExtId}` | `/stats/host-nodes/{hostExtId}` fallback |
| Protection policies | GET | `/datapolicies/{version}/config/protection-policies` | v4.2, v4.1; Data Protection v4.0 compatibility fallback |
| Recovery plans | GET | `/datapolicies/{version}/config/recovery-plans` | v4.2, v4.1; PC v3 `/nutanix/v3/recovery_plans/list` fallback |
| Recovery-plan stages and mappings | GET | `/datapolicies/{version}/config/recovery-plans/{planExtId}/{child}` | Child is `stages` or `network-mappings` |
| Recovery-plan execution history | GET/POST | `/dataprotection/{version}/config/recovery-plan-jobs` | v4.2, v4.1; PC v3 `/nutanix/v3/recovery_plan_jobs/list` fallback |
| PE Protection Domains | GET | `/PrismGateway/services/rest/v2.0/protection_domains/` | Direct PE, alternate PE v2 path, and PC proxy attempts |
| PE Protection Domain status | GET | `/PrismGateway/services/rest/v2.0/protection_domains/status` | Direct PE and PC proxy fallbacks |
| PE Remote Sites | GET | `/PrismGateway/services/rest/v2.0/remote_sites/` | Direct PE and PC proxy fallbacks |
| Subnets/VLANs | GET | `/networking/{version}/config/subnets` | v4.2, v4.1, v4.0.b1, v4.0 |
| Physical host NICs | GET | `/clustermgmt/{version}/config/clusters/{clusterExtId}/hosts/{hostExtId}/host-nics` | Additional Networking and PE v2 inventory fallbacks |
| Virtual switches, bonds, uplinks, bridges | GET | `/networking/{version}/config/{resource}` | Best-effort across v4.2, v4.1, v4.0.b1, v4.0 |
| Licensing | GET | `/licensing/{version}/config/{resource}` | v4.1, v4.0; resource is `entitlements`, `compliances`, or `violations` |
| LCM inventory | GET | `/lifecycle/{version}/resources/entities` | v4.2, v4.1, v4.0; read-only and filtered by cluster |
| NCC health checks | GET | `/opsmgmt/v4.0/monitoring/alerts` | Filtered to `HEALTH_CHECK` severity and cluster UUID |
| PE cluster CPU fallback | GET | `/PrismGateway/services/rest/v1/cluster/stats` | PE v2 stats paths and `/api/nutanix/v2.0` are also attempted |

## Version-selection behavior

- New primary paths use stable API releases newest-first. Cluster Management,
  Monitoring, and VMM currently try `v4.2`, `v4.1`, then `v4.0`.
- Existing Cluster Management and Networking collectors may still use
  `v4.0.b1` as a compatibility fallback until their individual v4 migrations
  are validated against all representative clusters.
- The script follows versioned self-links when Prism Central returns them.
- Legacy v3 calls remain where newer v4 responses may be empty even though the
  corresponding object is visible in Prism Central.
- Prism Element calls are fallbacks for cluster-specific data not consistently
  exposed by Prism Central.
- Lifecycle collection is read-only. The health check does not start LCM
  inventory, recommendation, or update operations.

## Authentication and transport

- Prism Central and Prism Element calls use HTTPS on port `9440` by default.
- The same supplied credentials are used for Prism Central and direct Prism
  Element fallback requests.
- TLS certificate verification is disabled to support environments using
  self-signed Prism certificates.
- The script uses read-only GET requests except for v3 list APIs, which require
  POST requests with list/filter bodies.
