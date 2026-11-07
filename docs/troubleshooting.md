# Troubleshooting

This page lists the most common problems of AppPackager with cause and fix. See [Error messages](reference/error-messages.md) for the messages.

## What it does

It gives a short fix for each common problem.

## When to use it

Use it when a check, a stage or a publish does not work.

## Requirements

None.

## Problems

| Message or symptom | Cause | Fix |
|---|---|---|
| **Check ConfigMgr**, **Publish to ConfigMgr** and **Convert ConfigMgr App to Intune** are disabled. | The console is not found on this computer. | Install the console. If it is installed in a custom path, set `SMS_ADMIN_UI_PATH` to its `bin\i386` folder and start the tool again. Hover over the button for the reason. |
| **Publish to Intune** or **Publish to WSUS** is disabled. | The system is not in use or not set up. | Set it in **Options**. See [Publish to Intune and Publish to WSUS](core/publish-intune-wsus.md). |
| A row says `Package error` and the log says `Network root path not accessible`. | The File Share Root is wrong or not reachable. | Correct it in **Options > ConfigMgr Preferences**. |
| **Check Latest** gives no version for GitHub applications. | The anonymous GitHub limit is used up. | Set `GITHUB_TOKEN`, or run `gh auth login`. |
| **Discover Unmanaged Software** says `Discovered 0 apps.` | No hidden packager matches the inventory, or the site has the applications already. | Read the status line. Check the inventory classes. |
| **One Click** shows `No apps are tracked yet.` | No application is tracked. | Track applications in **Options > One Click Settings**. |
| The CVE columns show `CPE missing`. | The packager has no CPE. | The row cannot be checked. |
| The window shows fewer than 312 applications. | The Product Filter hides some. | Open **Options > Product Filter**. |
| The Setup window does not show. | The preferences file exists. | Use **Options**. |

## Links

- [Error messages](reference/error-messages.md)
- [Getting started](getting-started.md)
