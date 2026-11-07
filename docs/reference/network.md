# Network

This page lists every outbound call of AppPackager.

## What it does

AppPackager has no service and no backend. Every call comes from one of the places in the table.

## When to use it

Use it to open the firewall for the tool, or to check the traffic.

## Requirements

| Who sends the call | What it sends | When |
|---|---|---|
| A packager script | A version lookup, then the installer download, with `curl.exe` against the vendor site. 92 packagers read `https://api.github.com/repos/<owner>/<repo>/releases/latest`. | Check Latest, One Click, Stage. |
| The GitHub token logic | The bearer token from `GITHUB_TOKEN`, `GH_TOKEN` or `gh auth token`. Without a token, the calls are anonymous: 60 for each hour and address. | Every GitHub call. |
| `start-apppackager.ps1` | The update check `https://api.github.com/repos/jasonulbright/app-packager/releases/latest`, once in 24 hours at the start. The icon pack and Content Prep downloads from Options. The `rate_limit` probe of the GitHub row. | At the start, and from Options. |
| Check CVEs | One NVD call for each checked application (`services.nvd.nist.gov`), EPSS from `api.first.org`, the CISA KEV catalog once a day (`www.cisa.gov`), and Shodan CVEDB (`cvedb.shodan.io`) when NVD fails. Open the firewall for all four names. | Check CVEs, CVE Report. |
| Check ConfigMgr, Publish to ConfigMgr | The console module to the SMS Provider, the file share. | Those actions. |
| Discover Unmanaged Software | One SQL connection to the site database, with Windows authentication and read-only intent. | Discover. |
| Publish to Intune, Convert with publishing | Microsoft Graph v1.0 and Azure Storage. `IntuneWinAppUtil.exe` is downloaded once from the Microsoft GitHub repository and its signature is checked. | Those actions. |
| Publish to WSUS | The WSUS API port of the server. | Publish to WSUS, WSUS Updates. |

The Application Workbench sends nothing.

## Steps

### See what one packager sends

Run the same command that the window runs:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File Packagers\package-notepadplusplus.ps1 -SiteCode XXX -GetLatestVersionOnly
```

## Options and settings of this page

None.

## Limits

- Behind a proxy that inspects TLS, `curl.exe` trusts the Windows certificate store. A failure names the URL and the HTTP status.
- A failed update check is logged and ignored. It never blocks the start.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| Check Latest fails for GitHub applications. | The anonymous GitHub limit is used up. | Set `GITHUB_TOKEN`, or run `gh auth login`. |
| A download fails. | The vendor site is blocked. | Open the URL in the log. |

## Links

- [Permissions](permissions.md)
- [Command line](command-line.md)
