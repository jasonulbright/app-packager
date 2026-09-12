# Review probe: prints only its own received arguments; does not install anything.
Write-Output (ConvertTo-Json -InputObject @($args) -Compress)
