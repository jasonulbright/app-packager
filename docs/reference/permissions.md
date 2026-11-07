# Permissions

This page lists the rights that each action of AppPackager needs.

## What it does

The tool runs with your Windows account. Configuration Manager, the file share, SQL Server, Intune and WSUS check each right.

## When to use it

Use it to ask for the right roles, or to find out why an action is refused.

## Requirements

A local administrator, to run the packager scripts.

## Steps

None.

## Options and settings of this page

The rights come from the README of the tool and the Microsoft documentation. The lab test of this guide used the lab administrator account, so it did not check a smaller role.

| Action | Right |
|---|---|
| Check Latest, Stage | Local administrator. Internet access to the vendor sites. |
| Check ConfigMgr | Read access to the applications of the site, and the console on this computer. |
| Publish to ConfigMgr | The right to create applications and deployment types. Write access to the File Share Root. The console. |
| Discover Unmanaged Software | `db_datareader` on the site database `CM_<site code>`. A domain account. |
| Publish to Intune | An Entra app registration with `DeviceManagementApps.ReadWrite.All`. |
| Publish to WSUS | Membership in the WSUS Administrators group, and access to the API port and the `UpdateServicesPackages` and `WSUSTemp` shares. |
| Convert ConfigMgr App to Intune | Read access to the application of the site. For publishing: the Intune right. |
| Script signing | Access to the private key of the certificate. A hardware key can need a PIN. |

## Limits

- The SQL views that Discover reads bypass the role-based administration of the console. Give `db_datareader` only to accounts that may read the inventory of all clients.

## Troubleshooting

| Message or symptom | Cause | Fix |
|---|---|---|
| Publish to ConfigMgr fails at the share. | No write access to the share. | Give the account write access. |
| Discover fails with an authentication message. | The account is not a domain account, or has no right on the database. | Use a domain account with `db_datareader`. |

## Links

- [Network](network.md)
- [External references](../external-references.md)
