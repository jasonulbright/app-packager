# Files and logs

This page lists every folder and file that AppPackager reads or writes.

## What it does

The tool writes logs and reports in the install folder, the staged content in the download root, and its history in your user profile.

## When to use it

Use it to find a log, to clean up or to check what the tool leaves on a computer.

## Requirements

None.

## Steps

None.

## Options and settings of this page

| Path | Content |
|---|---|
| `Logs\` | One log for each packager run, for example `package-7zip-stage-<date>-<time>.out.log`. The window shows the same lines. The log keeps the full text. |
| `Logs\Reports\cve-report-<timestamp>.html`, `...-applications.csv`, `...-cves.csv`, and the workbook, Markdown and JSON files | The CVE report. |
| `Logs\OneClick\one-click-<timestamp>.md`, `.json` | The report of each One Click run. |
| `<Download Root>\<App>\<Version>\` | The staged content of an application, with `stage-manifest.json`. |
| `<File Share Root>\Applications\<Vendor>\<Application>\<Version>\` | The packaged content on the share, for the default layout. |
| `<Output folder>\ConfigMgrImport\<application>\` | The plan, the report and the content of a conversion to Intune. |
| `%LOCALAPPDATA%\AppPackager\app-history.json` | The history of the checks. |
| `%LOCALAPPDATA%\AppPackager\update-check.json` | The result of the daily update check. |
| `%LOCALAPPDATA%\AppPackager\vuln-results\` | The CVE results, one file for each CPE and version pair, used for six hours. |
| `%LOCALAPPDATA%\AppPackager\vuln-products\` | The NVD records of each product, used for six hours. |
| `%LOCALAPPDATA%\AppPackagerData\Workbench\` | The Workbench profiles and the build records. |

## Limits

- The tool never deletes staged content. Delete old versions by hand.
- A log file stays open while its packager runs.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| A log file cannot be deleted. | A packager of an earlier run still runs. | End that process. |
| The disk of the download root is full. | Staged versions pile up. | Delete the old version folders. |

## Links

- [Settings](settings.md)
- [Stage and Publish to ConfigMgr](../core/stage-publish.md)
