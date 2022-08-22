# Under Development
## Objectives
- [x] Fix known bugs in existing 2.6 version
- [ ] Add reporting for the following use cases
  - [x] FC Storage Ports
  - [ ] Standard and FCOE Appliance Ports
  - [ ] Port Channels
  - [ ] Disjoint L2
- [x] Separate HTML report template (including CSS and JScript) to ease updates
- [ ] Improve performance
- [x] Change minimum PowerShell version to 5 - **Complete**
- [ ] Apply Python style coding practices according to [The PowerShell Style Guide](https://github.com/PoshCode/PowerShellPracticeAndStyle/blob/master/Style-Guide/Introduction.md)
- [ ] Refactor where possible to make code more DRY and readable

This version will be "UCS Configuration Report" and no longer "UCS Health Check" and will diverge from the originally produced Cisco version.

## UCS Health Check and Inventory Script (ported from UCS Communities)
See original [README](https://github.com/datacenter/ucs-browser)

## Procedures
### PowerShell Configuration
1. Verify PowerShell 5 or higher is installed on the management station (`$PSVersionTable`)
   1. **Windows**
      1. Issue command: `$PSVersionTable`
      2. If below 5, then install latest management pack
   2. **Mac**
      1. Review [Microsoft KB Article][1] for prerequisites.
2. Install VMware PowerCLI tools (Check with `Get-Module -ListAvailable -Name Cisco.UCSManager | select Name,Version`)
   1. **Windows (Internet Required) (Run as Administrator Required)**:
      1. Update PowerShell Help
         1. `Update-Help -Force -ErrorAction SilentlyContinue`
      2. Update NuGet package provider
         1. `Install-PackageProvider NuGet -Force -Verbose`
      3. Force update of PowerShellGet. UCS Power Tools install will identify this as being out of date.
         1. `Install-Module PowerShellGet -Force -AllowClobber`
      4. Restart PowerShell (Run as Administrator Required)
      5. Install UCS Power Tools from Internet Repository (NuGet)
         1. `Install-Module -Name Cisco-UCSManager`
      6. Lower PowerShell Execution Policy level (Check with company security policy)
         1. `Set-ExecutionPolicy -Unrestricted`
      7. Verify new module has been properly loaded.
         1. `Import-Module Cisco.UCSManager`
   2. **Mac (Internet Required)**:
      1. Update PowerShell Help
         1. `Update-Help -Force -ErrorAction SilentlyContinue`
      2. Install UCS Power Tools from Internet Repository (NuGet)
         1. `Install-Module VMware.PowerCLI`
      3. Verify new module has been properly loaded.
         1. `Import-Module VMware.PowerCLI`

[1]: https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-macos?view=powershell-7.2

[2]: https://kb.vmware.com/s/article/59235
