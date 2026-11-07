# Check CVEs and the CVE report

Check CVEs looks up the vulnerabilities that an update from the version on your site to the latest version fixes.

## What it does

For each checked application, the tool queries the National Vulnerability Database (NVD) with the CPE name of the packager and the version range. A CVE that affects the version on the site and not the latest version counts as **fixed by updating**. A CVE that affects both counts as **still in latest**. The tool adds the EPSS score and the flag of the Known Exploited Vulnerabilities (KEV) catalog.

## When to use it

Use it to decide which updates are urgent.

## Requirements

- A `CPE` header in the packager, a version from Check ConfigMgr, and a latest version from Check Latest. A row that lacks one shows the reason in the **CVEs** column: `CPE missing`, `No ConfigMgr version` or `No latest version`.
- Access to `services.nvd.nist.gov`, `api.first.org` and `www.cisa.gov`. See [Network](../reference/network.md).

## Steps

### Check the CVEs

1. Check the rows. Run [Check Latest and Check ConfigMgr](versions.md) first.
2. Select **Check CVEs**. A progress overlay shows.

![The window while Check CVEs runs](../images/cve-running.png)

3. Wait for the status `Complete.` The columns **CVEs**, **Max CVSS**, **KEV**, **In latest** and **Newest CVE** fill in. The log shows a line for each application, for example `0 CVEs between 26.03 and 26.06 (NVD)`.

![The window after Check CVEs](../images/cve-done.png)

The CVE count in the grid opens a window with every CVE and its links.

### Write the CVE report

1. Check the rows. With no row checked, the report covers all rows.
2. Select **CVE Report**. The tool writes these files in the folder `Logs\Reports`: an HTML page, two CSV files, a workbook, a Markdown file and a JSON file. Each name starts with `cve-report-<timestamp>`. The status bar says `CVE report written`.

![The window after CVE Report](../images/cve-report-written.png)

3. The HTML page opens in the browser.

![The HTML page of the CVE report](../images/cve-report-browser.png)

## Options and settings of this page

| Setting (Options > ConfigMgr Preferences) | Effect |
|---|---|
| NVD API key | Raises the NVD limit from 5 to 50 requests in 30 seconds. The variable `NVD_API_KEY` does the same. |
| Look up CVEs after Check ConfigMgr | Chains Check CVEs after Check ConfigMgr. |

## Limits

- NVD allows 5 requests in 30 seconds without a key. A large run is slow without a key.
- The tool caches a result for six hours for each CPE and version pair, in `%LOCALAPPDATA%\AppPackager\vuln-cache.json`.
- When NVD fails, the tool asks Shodan CVEDB with the version on the site. Its records carry no version ranges, so every record counts as affecting the installed version.
- The lab application in this guide had no CVE between the two versions. The CVE window and a row with CVEs were not captured.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| `CPE missing` | The packager has no `CPE` header. | The row cannot be checked. |
| `No ConfigMgr version` | Check ConfigMgr did not run for the row. | Run Check ConfigMgr. |
| The run is slow and repeats. | NVD throttles calls without a key. | Set an NVD API key. |

## Links

- [Check Latest and Check ConfigMgr](versions.md)
- [Network](../reference/network.md)
