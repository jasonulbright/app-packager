# AppPackager documentation

AppPackager packages 312 enterprise applications for Configuration Manager, Intune and WSUS. It checks the vendor for the latest version, downloads and verifies the installer, writes the install and uninstall scripts and the detection rule, and creates the application. This page links every page of the documentation.

## Start

- [Getting started](getting-started.md): install, the Setup window and the first look at the window.

## Core

- [Main window](core/main-window.md): the sidebar, the grid, the filter and the log.
- [Check Latest and Check ConfigMgr](core/versions.md): read the vendor version and the version that the site has.
- [Stage and Publish to ConfigMgr](core/stage-publish.md): build the content locally, then create the application on the site.
- [Publish to Intune and Publish to WSUS](core/publish-intune-wsus.md): the other two destinations.
- [Discover Unmanaged Software](core/discover.md): find software that clients run and the site does not manage.
- [Check CVEs and the CVE report](core/cve.md): the vulnerabilities of the version on the site, and the ones that an update fixes.
- [One Click](core/one-click.md): the unattended run of the apps that you track.
- [Application Workbench](core/workbench.md): change how one application is built, without editing its script.
- [Add Installer](core/add-installer.md): package an installer that you drop on the window.
- [Convert ConfigMgr App to Intune](core/convert-to-intune.md): make Intune apps from an existing Configuration Manager application.
- [Options dialog](core/options.md): every page of the settings.

## Reference

- [Settings](reference/settings.md): what the tool saves and where.
- [Files and logs](reference/files-and-logs.md): every folder and file that the tool writes.
- [Permissions](reference/permissions.md): the rights that each action needs.
- [Network](reference/network.md): every outbound call.
- [Command line](reference/command-line.md): run a packager or a build without the window.
- [Error messages](reference/error-messages.md): message, cause and fix.
- [Glossary](reference/glossary.md): the terms of the tool.
- [Capture the images again](reference/capture-images.md): the one command that regenerates the images of this guide.

## External references

- [External references](external-references.md): Microsoft pages for each destination.

## Troubleshooting

- [Troubleshooting](troubleshooting.md): the most common problems with cause and fix.
