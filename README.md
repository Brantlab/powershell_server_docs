# Server Inventory To Markdown

This PowerShell script walks an AD OU, collects useful server details, and writes one Markdown file per server for MKDocs.

The Markdown is tuned for Material for MkDocs with admonitions, card grids, and collapsible detail sections.

## What It Collects

- Server name and DNS name
- Operating system and version
- Latest hotfix seen plus the 10 most recent hotfixes
- Last boot time and uptime
- Disk size and free space for every local disk
- Memory, model, serial number, and basic network details
- Reachability and collection errors so gaps are obvious in the output

## Files

- `Generate-ServerInventoryDocs.ps1`: main report generator
- `Update-InventoryRepo.ps1`: optional wrapper to generate docs and commit them into Git
- `docs/servers/*.md`: generated output location by default
- `docs/servers/SAMPLE-SRV01.md`: sample output you can preview in MkDocs right away

## Example Usage

```powershell
$credential = Get-Credential

.\Generate-ServerInventoryDocs.ps1 `
  -SearchBase "OU=Servers,OU=Production,DC=contoso,DC=com" `
  -DomainController "dc01.contoso.com" `
  -OutputPath ".\docs\servers" `
  -Credential $credential
```

If the account running the script already has access to AD and remote WMI/CIM on the target servers, you can omit `-Credential`.

## Overnight Run Pattern

Use Task Scheduler, a CI runner on a Windows host, or an existing automation server to run the script nightly. A common pattern is:

1. Pull the Git repo onto a Windows machine that can reach AD and the servers.
2. Run `Generate-ServerInventoryDocs.ps1`.
3. Commit the changed Markdown files.
4. Push to the repo that triggers your MKDocs runner.

If you want one command to handle generation plus Git commit:

```powershell
$credential = Get-Credential

.\Update-InventoryRepo.ps1 `
  -SearchBase "OU=Servers,OU=Production,DC=contoso,DC=com" `
  -DomainController "dc01.contoso.com" `
  -OutputPath ".\docs\servers" `
  -Credential $credential `
  -Push
```

## Notes

- The script expects the `ActiveDirectory` module to be available.
- Remote server collection uses CIM queries. Make sure firewall and permissions allow that from the runner.
- Unreachable servers still get a Markdown page so the absence of data is visible.
- For the intended look in Material for MkDocs, enable the usual extensions for admonitions/details and markdown-in-HTML styling.
