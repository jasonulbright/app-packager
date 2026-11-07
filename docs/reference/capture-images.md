# Capture the images again

One command regenerates every image of this guide.

## What it does

`tools\guide-capture\Invoke-GuideCapture.ps1` copies the tool to a lab computer, starts it in an interactive session of a lab account, walks the windows and dialogs with Windows UI Automation, takes the screenshots, copies the PNG files back to `docs\images` and removes its files from the lab computer. It replaces only the images that exist already. With `-AddNew` it also adds new images.

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

A second run does the missing-console test on a copy of the tool whose console detection reports `not found`. The command removes the saved settings and the staged files of the lab account before and after each run.

## Limits

- The capture needs an interactive session on the lab computer. A locked session breaks it.
- The images show the data of the lab site.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| `window not found` | The window did not open in time. | Run again. |
| The command cannot delete a log of the lab copy. | A packager of an earlier run still runs. | The command ends it at the next start. Run again. |

## Links

- [Index](../index.md)
