# Cisco UCS Configuration Report
## Synopsis
This script generates an HTML report showing all key configuration points on a Cisco UCS domain that is managed by UCS Manager (vs. Cisco Intersight).

## Acknowledgement
This project was forked from code written by Brandon Beck (Cisco) and ported to [GitHub][4] from an original [Cisco Community Post][5].

## Prerequisites
This script requires both:
- PowerShell 5 or later
  - Check version with command: `$PSVersionTable`
  - For systems with older PS version, upgrade the Windows Management Framework
- Cisco UCS PowerTool Suite 3 or later
  - If installed, check version with command: `Get-Module -ListAvailable -Name Cisco.UCSManager | select Name,Version`
  - Available via PowerShell Gallery or direct download in PS using NuGet package provider

**Apple Mac:** Review [Microsoft KB Article][1] for installation of PowerShell.

**Linux:** Review [Microsoft KB Article][6] for installation of PowerShell.

## Git Branches
The current stable version will always be in the master branch.

## Procedures
### Setup
1. Install UCS Power Tool (Check with `Get-Module -ListAvailable -Name Cisco.UCSManager | select Name,Version`)
   1. **Windows (Internet Required) (Run as Administrator Required)**:
      1. Update PowerShell Help
         1. `Update-Help -Force -ErrorAction SilentlyContinue`
      2. Update NuGet package provider
         1. `Install-PackageProvider NuGet -Force -Verbose`
      3. Force update of PowerShellGet. UCS Power Tools install will identify this as being out of date.
         1. `Install-Module PowerShellGet -Force -AllowClobber`
      4. Restart PowerShell (Run as Administrator Required)
      5. Install UCS Power Tools from Internet Repository (NuGet)
         1. `Install-Module -Name Cisco.UCSManager`
      6. Lower PowerShell Execution Policy level (Check with company security policy)
         1. `Set-ExecutionPolicy -Unrestricted`
      7. Verify new module has been properly loaded.
         1. `Import-Module Cisco.UCSManager`
   2. **Mac / Linux (Internet Required)**:
      1. Update PowerShell Help
         1. `Update-Help -Force -ErrorAction SilentlyContinue`
      2. Install UCS Power Tools from Internet Repository (NuGet)
         1. `Install-Module Cisco.UCSManager`
      3. Verify new module has been properly loaded.
         1. `Import-Module Cisco.UCSManager`
2. Download Cisco UCS Configuration Report project from GitHub (your choice)
   - Direct browser download
   - Git CLI: `git clone https://github.com/pschapman/ucs-config-report.git`

**NOTE:** UCS_Config_Report.ps1 and Report_Template.htm are mandatory files and must be in the same directory.

### Creating Reports
#### Manual (All platforms as of v4.1)
1. At a PowerShell prompt run `UCS_Config_Report.ps1`
2. Select option 1 on the Main Menu to manage domain connections
3. Select option 1 on the Connection Management Menu to connect to a UCS domain
   1. Input IP/FQDN and credentials as prompted
4. Select option 7 on the Connection Management Menu to return to the Main Menu
5. Select option 2 on the Main Menu to create a report
   1. Select a location and file name for the report
6. Select option Q on the Main Menu to exit the program (automatically disconnects from UCS domain)

#### Automated
1. Create a Credential Cache File
   1. At a PowerShell prompt run `UCS_Config_Report.ps1`
   2. Select option 1 on the Main Menu to manage domain connections
   3. Select option 1 on the Connection Management Menu to connect to a UCS domain
      1. Input IP/FQDN and credentials as prompted
   4. Select option 3 on the Connection Management Menu to save credentials
   5. Select option 7 on the Connection Management Menu to return to the Main Menu
   6. Select option Q on the Main Menu to exit the program
2. At the PowerShell prompt run the script with arguments
   1. `./UCS_Config_Report.ps1 -UseCached -RunReport -Silent`
3. (Optional) Configure Scheduled Task (Windows) or cron job (Mac) to run the script on a regular basis

## What's New
**Version 4.1 - 8/28/2022**
- New Features
  - Tested on both linux and Mac
    - Credentials file is no longer mandatory to run report. Since windows dialog is unavailable, the script will drop the report in the script directory named, "UCS_Report_YYYY_MM_DD_hh_mm_ss.html"
  - Fixed population of temperature data for rack mounts (e.g., HyperFlex) and updated column headers
  - Term "health check" replaced throughout with "config report" or similar.
  - Duplicate DNS and NTP servers no longer listed.
  - "No Stats" execution. Still experimental. Added to improve performance when testing.
- Code Revisions
  - Multiple ScriptBlocks converted to functions to allow for test isolation
    - Conversion required script to be run recursively.  Added needed hidden CLI options for operation.
    - Global script statements moved to "Main" function
  - Found and fixed all errors thrown by main data gathering process (but hidden by opaque execution)
  - Functions renamed using standard (verb)-(action) style

**Version 4.0 - 8/22/2022**
- New Features
  - Empty memory slots now show as "empty" instead of "undefined/indeterminate". [Screenshot](DocImages/EmptyMemorySlots.png)
  - FC Storage Ports, FCOE Ports, and Appliance Ports now properly reported under the SAN tab. Added column to distinguish interface roles (e.g., network, server, storage). [Screenshot](DocImages/NewSANReporting.png)
  - File Save dialog box (Windows) now defaults to the Desktop folder for output and offers a pre-created file name. [Screenshot](DocImages/FileSaveDialog.png)
    - Silent operation will save report files to script folder
- Code Revisions
  - Fixed Issues 2, 3, & 4 in the [parent code base][4]
  - Fixed discovered issues [#1][i001], [#2][i002], [#3][i003], [#4][i004], and [#5][i005] found during initial code cleanup
  - Switched functions and script to Documentation Strings for PS
  - Linted with VSCode and resolved all reported issues.

## Update Objectives in Progress
- [ ] Improve performance
- [ ] Apply Python style coding practices according to [The PowerShell Style Guide][3]
- [ ] Refactor where possible to make code more DRY and readable
- [ ] Add reporting for the following use cases
  - [ ] Port Channels
  - [ ] Disjoint L2

[1]: https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-macos?view=powershell-7.2
[2]: https://kb.vmware.com/s/article/59235
[3]: https://github.com/PoshCode/PowerShellPracticeAndStyle/blob/master/Style-Guide/Introduction.md
[4]: https://github.com/datacenter/ucs-browser
[5]: https://community.cisco.com/t5/unified-computing-system-knowledge-base/ucs-healthcheck-v2-5/ta-p/3654629
[6]: https://docs.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-linux?view=powershell-7.2
[i001]: https://github.com/pschapman/ucs-config-report/issues/1
[i002]: https://github.com/pschapman/ucs-config-report/issues/2
[i003]: https://github.com/pschapman/ucs-config-report/issues/3
[i004]: https://github.com/pschapman/ucs-config-report/issues/4
[i005]: https://github.com/pschapman/ucs-config-report/issues/5
