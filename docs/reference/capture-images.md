# Capture the images again

One command regenerates every image of this guide.

## What it does

`tools\guide-capture\Invoke-GuideCapture.ps1` copies the tool to a lab computer, starts it in an interactive session of a lab account, walks the windows and dialogs with Windows UI Automation, takes the screenshots, copies the PNG files back to `docs\images` and removes its files from the lab computer. It replaces only the images that exist already. With `-AddNew` it also adds new images.

The command also takes the README images `main-dark.png`, `setup.png`, `workbench.png` and `drop-preview.png` from the same run, in the `screenshots` folder. The README images `one-click-plan.png` and `wsus-updates.png` have no capture step: the One Click plan needs a configured site with tracked applications, and the lab has no WSUS. They keep their old image.

## When to use it

Use it after a change that alters a window, and before a release.

## Requirements

- A lab computer with the Configuration Manager console, an SMS Provider that you can read, access to the vendor sites and to GitHub, and PowerShell remoting for an account.
- A script, outside the repository, that sets `$LabCredential`.
- A domain account that has an interactive session on the lab computer.
- Lab data only. The capture reads the site and downloads one installer to a local folder of the lab computer. The File Share Root of the capture points to a server that does not exist, so Publish to ConfigMgr stops at its share check. The capture never publishes to Intune or WSUS, never saves a Workbench profile and never starts a One Click run.

## Steps

1. Run this command from the repository folder:

```powershell
.\tools\guide-capture\Invoke-GuideCapture.ps1 -CredentialScript C:\lab\credentials.local.ps1 -Target CLIENT01 -LogonUser contoso\labadmin -SmsProvider cm01.contoso.com
```

2. Wait for the command to end. It takes about 12 minutes. It prints a warning for each step that failed and keeps the captured files for you to read.
3. Open every changed image and check it.
4. Run the documentation test: `Invoke-Pester .\Tests\Documentation.Tests.ps1`.

The Chrome CVE images (`cve-chrome-*.png` and `light-cve-chrome-*.png`) are not part of this run. Make them with the switch `-OnlyCveChrome`, and add `-AddNew` the first time. The run writes a packager history with Chrome 157.0.8103.40 on the site and 158.0.8127.37 as the latest version, checks only the Chrome row, and runs Check CVEs against the live NVD. It needs access to `services.nvd.nist.gov`. Allow `-Minutes 45`. Open the images and check that the CVE window shows a CVE with CVSS 9.6 and the status `Fixed`. If NVD no longer lists such a CVE for these versions, the step stops with an error.

A second run does the missing-console test on a copy of the tool whose console detection reports `not found`. The command removes the saved settings of the lab account and the staging folder `C:\guide\app-packager\dl` before and after each run, so that an earlier stage cannot be reused.

## Limits

- The capture needs an interactive session on the lab computer. A locked session breaks it.
- One capture runs at a time on a lab computer. The command takes the lock file `C:\guide\capture.lock` and waits while another capture holds it. A lock that nobody touched for 60 minutes is stale. The command works in its own folder `C:\guide\app-packager`, removes only its own files and the lock, and never removes `C:\guide`.
- The images show the data of the lab site, except the Chrome CVE images.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| `window not found` | The window did not open in time. | Run again. |
| The command cannot delete a log of the lab copy. | A packager of an earlier run still runs. | The command ends it at the next start. Run again. |

## Links

- [Index](../index.md)
