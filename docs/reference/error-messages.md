# Error messages

This page lists the messages of AppPackager with cause and fix.

## What it does

It explains each message that this guide met, and what to do.

## When to use it

Use it when the tool shows a message that you do not know.

## Requirements

None.

## Steps

Find the message in the table. Do the fix.

## Options and settings of this page

| Message | Cause | Fix |
|---|---|---|
| `Network root path not accessible: <path>` | The File Share Root does not exist, or the account has no access. | Correct the path in Options. Check the rights. |
| `Package error` (status of a row) | The Package phase failed. | Read the log. |
| A disabled **Check ConfigMgr**, **Publish to ConfigMgr** or **Convert ConfigMgr App to Intune** | The console is not found. | Install the console, or set `SMS_ADMIN_UI_PATH`. |
| `No apps are tracked yet.` (message `One Click Not Configured`) | One Click has no tracked application. | Track applications in Options, Updates and CVEs, One Click. |
| `CPE missing`, `No version` | A CVE lookup lacks an input. | Add the missing input. See [Check CVEs](../core/cve.md). |
| `Not in NVD` | NVD holds no record under the product name of the `CPE` header. | Compare the header with the NVD CPE dictionary. See [Check CVEs](../core/cve.md#a-product-that-nvd-does-not-know). |
| `WSUS: not supported` | WSUS cannot carry the application. | Read the reason in the log. |
| `Windows authentication to SQL Server <server> failed.` | Discover cannot sign in. | Use a domain account on a domain-joined computer. |

## Limits

The table lists the messages that the code and this guide show. A message of a packager shows as the packager writes it.

## Troubleshooting

See [Troubleshooting](../troubleshooting.md).

## Links

- [Troubleshooting](../troubleshooting.md)
