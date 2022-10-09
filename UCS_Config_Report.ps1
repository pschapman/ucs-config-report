#Requires -Version 5

<#
.SYNOPSIS
    Generates a UCS Configuration Report

.DESCRIPTION
Script for polling multiple UCS Domains and creating an html report of the domain inventory, configuration, and
port statistics.  Internet access is required to view report due to CDN content pull for CSS and JavaScript.

.PARAMETER UseCached
Run script using 'ucs_cache.ucs' file to login to all cached UCS domains.

.PARAMETER RunReport
Go directly to generating a configuration report. Prompts for report file name unless -Silent also included.

.PARAMETER Silent
Bypasses prompts and menus. Auto-names report as [UCS_Report_YY_MM_DD_hh_mm_ss.html]. Should be used with
-RunReport and -UseCached.

.PARAMETER Email
Email the report file after automated execution.  Must specify target email address. (e.g., user@domain.tld)

.PARAMETER NoStats
    Skip statistics collection --WARNING-- Experimental feature.

.EXAMPLE
UCS_Config_Report.ps1 -UseCached -RunReport -Silent -Email user@domain.tld

Run Script with no interaction and email report. UCS cache file must be populated first. Can be run as a scheduled
task.

.NOTES
Script Version: 4.3
JSON Schema Version 4.2
Attributions::
    Author: Paul S. Chapman (pchapman@convergeone.com) 09/03/2022
    History: UCS Configuration Report forked from UCS Health Check v2.6
    Source: Brandon Beck (robbeck@cisco.com) 05/11/2014
    Contribution: Marcello Turano

Before generating the report select option (1) from the root menu and connect to target UCS Domains. Cache
domain credentials after connecting to speed up future script executions.

.LINK
https://github.com/pschapman/ucs-config-report
#>

Param(
    [Switch] $UseCached,
    [Switch] $RunReport,
    [Switch] $Silent,
    [String] $Email = $null,
    [Switch] $NoStats,
    [Parameter(DontShow)][Switch] $RunAsProcess,
    [Parameter(DontShow)] $Domain,
    [Parameter(DontShow)] $ProcessHash
)

# Global Variable Definition
# ==========================
$UCS = @{}                              # Hash variable for storing UCS handles
$UCS_Creds = @{}                        # Hash variable for storing UCS domain credentials
$runspaces = $null                      # Runspace pool for simultaneous code execution
$dflt_output_path = [System.Environment]::GetFolderPath('Desktop') # Default output path. Alternate: $pwd

# Determine script platform
# ==========================
If ($PSVersionTable.PSVersion.Major -ge '6') {$platform = $PSVersionTable.Platform} else {$platform = 'Windows'}

# Email Variables
# =========================================================
$mail_server = ""        # Example: "mxa-00239201.gslb.pphosted.com"
$mail_from = ""          # Example: "Cisco UCS Configuration Report<user@domain.tld>"
# =========================================================
$test_mail_flag = $false # Boolean - Enable for testing without CLI argument
$test_mail_to = ""       # Example: "user@domain.tld"

function Start-Main {
    <#
    .DESCRIPTION
        Run pre-checks and go to menus or silent operation
    #>
    # Ensures cached credentials are passed for automated execution
    if($UseCached -eq $false -and $RunReport -eq $true) {
        Write-Host "`nCached Credentials must be specified to run report (-UseCached)`n`n"
        exit
    }

    # Check that required modules are present
    Test-RequiredPsModules

    # Run data gather in parallel iteration of script for asynchronous processing of multiple UCS domains
    if ($RunAsProcess) {
        Invoke-UcsDataGather -domain $Domain -Process_Hash $ProcessHash
        exit
    }

    # Loads cached ucs credentials from current directory
    if($UseCached) {
        If(Test-Path "$((Get-Location).Path)\ucs_cache.ucs") {
            Add-UcsHandleAndCreds
        } else {
            Write-Host "`nCache File not found at $((Get-Location).Path)\ucs_cache.ucs`n`n"
            exit
        }
    }

    # Automates configuration report execution if RunReport switch is passed
    If($UseCached -and $RunReport) {
        Start-UcsDataGather
        if($Silent) {
            Disconnect-AllUcsDomains
            exit
        }
    }

    # Start the Main Menu
    Show-MainMenu
}

function Show-MainMenu {
    <#
    .DESCRIPTION
        Text driven menu interface for connecting to UCS domains or running reports.
    #>
    # Main Menu
    $main_menu = "
            MAIN MENU

        1. Connect/Disconnect UCS Domains
        2. Generate UCS Configuration Report
        Q. Exit Program
    "
    :menu
    while ($true) {
        Clear-Host
        Write-Host $main_menu
        $command = Read-Host "Enter Command Number"
        Switch ($Command) {
            # Connect to UCS domains
            1 {Show-CnxnMgmtMenu}

            # Run UCS Configuration Report
            2 {Start-UcsDataGather}

            # Cleanly exit program
            'q' {
                Disconnect-AllUcsDomains
                break menu
            }
        }
    }
}

function Test-UcsHandle ($Domain) {
    <#
    .DESCRIPTION
        Checks if the passed hash variable contains an active UCS handle.
    .PARAMETER Domain
        Hash table containing domain VIP, credentials, and handle.
    .OUTPUTS
        [bool] - Existance of handle
    #>
    $error.clear()
    try {Get-UcsStatus -Ucs $Domain.Handle | Out-Null}
    catch {
        return $false
    }
    if (!$error) {return $true}
}

function Confirm-AnyUcsHandle {
    <#
    .DESCRIPTION
        Checks if any UCS handle exists in the global UCS hash variable
    .OUTPUTS
        [bool] - Existance of handle
    #>
    foreach ($Domain in $UCS.get_keys()) {
        if(Test-UcsHandle($UCS[$Domain])) {return $true}
    }
    return $false
}

function Add-UcsHandleAndCreds {
    <#
    .DESCRIPTION
        Connects to a ucs domain either by interactive user prompts or using cached credentials if the UseCached
        switch parameter is passed
    .OUTPUTS
        [none] Updates script global $UCS and $UCS_Creds hash tables
    #>
    # If UseCached parameter is passed then grab all UCS credentials from cache file and attempt to login
    if($UseCached) {
        Clear-Host
        # Grab each line from the cache file, remove all white space, and pass to a foreach loop
        Get-Content "$((Get-Location).Path)\ucs_cache.ucs" | Where-Object {$_.trim() -ne ""} | ForEach-Object {

            # Split credential data - each line consists of UCS VIP, username, hashed password
            $credData = $_.Split(",")

            # Ensure we have all three components if the credential data
            if($credData.Count -eq 3) {
                # Clear system $error variable before trying a UCS connection
                $error.clear()
                try {
                    # Attempts to create a UCS handle and stores the handle into the global UCS hash variable if connection is successful
                    $domain = @{}
                    $domain.VIP = $credData[0]
                    # Creates a credential variable from the username and hashed password pulled from the cache entry
                    $domain.Creds = New-Object System.Management.Automation.PsCredential($credData[1], ($credData[2] | ConvertTo-SecureString))
                    $domain.Handle = Connect-Ucs $domain.VIP -Credential $domain.Creds -NotDefault -ErrorAction SilentlyContinue
                    # Checks that handle actually exists
                    Get-UcsStatus -Ucs $domain.Handle | Out-Null
                }
                # Catch any failed domain connections
                catch [Exception] {
                    # Allow user to continue/exit script execution if a connection fails
                    $ans = Read-Host "Error connecting to UCS Domain at $($domain.VIP)  Press C to continue or any other key to exit"
                    Switch -regex ($ans.ToUpper()) {
                        "^[C]" {continue}
                        default {exit}
                    }
                }
                # Display a message to the user that the attempted UCS domain connection was successful and add handle to global UCS variable
                if (!$error) {
                    Write-Host "Successfully Connected to UCS Domain: $($domain.Handle.Ucs)"
                    $domain.Name = $domain.Handle.Ucs
                    $script:UCS.Add($domain.Handle.Ucs, $domain)
                    $script:UCS_Creds[$domain.Handle.Ucs] = @{}
                    $script:UCS_Creds[$domain.Handle.Ucs].VIP = $domain.VIP
                    $script:UCS_Creds[$domain.Handle.Ucs].Creds = $domain.Creds
                }
            }
        }
        Start-Sleep(1)
    } else {
    # Connect to a single UCS domain through an interactive prompt
        while ($true) {
            Clear-Host
            # Prompts user for UCS IP/DNS and user creential
            Write-Host "Please enter the UCS Domain information"
            $domain = @{}
            $domain.VIP = Read-Host "VIP or DNS"
            Write-Host "Prompting for username and password..."
            $domain.Creds = Get-Credential

            # Clear error variable to check for failed connections
            $error.clear()
            try {
                # Attempt UCS connection from entered data
                $domain.Handle = Connect-Ucs $domain.VIP -Credential $domain.Creds -NotDefault -ErrorAction SilentlyContinue
                # Checks that handle actually exists
                Get-UcsStatus -Ucs $domain.Handle | Out-Null

            }
            # Catch failed connection attempts and allow user to re-enter credentials
            catch [Exception] {
                # Press M to return to menu or any other key to re-enter credentials
                $ans = Read-Host "Error connecting to UCS Domain.  Press enter to retry or M to return to Main Menu"
                Switch -regex ($ans.ToUpper()) {
                    "^[M]" {return}
                    default {continue}
                }
            }
            if (!$error) {
                # Notify the user that the connection was successful and add handle to global UCS variable
                Write-Host "`nSuccessfully Connected to UCS Domain: $($domain.Handle.Ucs)"
                Write-Host "Redirecting to Main Menu..."
                $domain.Name = $domain.Handle.Ucs
                $script:UCS.Add($domain.Handle.Ucs, $domain)
                $script:UCS_Creds[$domain.Handle.Ucs] = @{}
                $script:UCS_Creds[$domain.Handle.Ucs].VIP = $domain.VIP
                $script:UCS_Creds[$domain.Handle.Ucs].Creds = $domain.Creds
                Start-Sleep(2)
                break
            }
        }
    }
}

function Show-CnxnMgmtMenu {
    <#
    .DESCRIPTION
        Text driven menu interface for allowing users to connect, disconnect, and cache UCS domain information
    #>
    $conn_menu = "
       CONNECTION MANAGEMENT

    1. Connect to a UCS Domain
    2. List Active Sessions
    3. Cache current connections
    4. Clear session cache
    5. Select Session for Disconnect
    6. Disconnect all Active Sessions
    7. Return to Main Menu`n"

    while ($true) {
        Clear-Host
        Write-Host $conn_menu
        $option = Read-Host "Enter Command Number"
        Switch ($option) {
            1 {Add-UcsHandleAndCreds}

            # Print all active UCS handles to the screen
            2 {
                Clear-Host
                if(!(Confirm-AnyUcsHandle)) {
                    Read-Host "There are currently no connected UCS domains`n`nPress any key to continue"
                    break
                }
                $index = 1
                Write-Host "`t`tActive Session List`n$("-"*60)"
                foreach ($Domain in $UCS.get_keys()) {
                    # Checks if the UCS domain is active and prints a formatted list
                    if(Test-UcsHandle($UCS[$Domain])) {
                        # Composite format using -f method
                        "$($index)) {0,-28} {1,20}" -f $($UCS[$Domain].Name),$UCS[$Domain].VIP
                        $index++
                    }
                }
                Read-Host "`nPress any key to return to menu"
            }

            # Cache all UCS handles to a cache file for future reference
            3 {
                # Check for an empty domain list
                if($UCS_Creds.Count -eq 0) {
                    Read-Host "`nThere are currently no connected UCS domains`n`nPress any key to continue"
                    break
                }
                # Iterate through UCS domain hash and store information to cache file
                foreach ($Domain in $UCS.get_keys()) {
                    # If the cache file already exists remove any lines that match the current domain name
                    If(Test-Path "$((Get-Location).Path)\ucs_cache.ucs") {
                        (Get-Content "$((Get-Location).Path)\ucs_cache.ucs") | ForEach-Object {$_ -replace "$($UCS_Creds[$Domain].VIP).*", ""} | Set-Content "$((Get-Location).Path)\ucs_cache.ucs"
                    }
                    # Add the current domain access data to the cache
                    $UCS_Creds[$Domain].VIP + ',' + $UCS_Creds[$Domain].Creds.Username + ',' + ($UCS_Creds[$Domain].Creds.Password | ConvertFrom-SecureString) | Add-Content "$((Get-Location).Path)\ucs_cache.ucs"
                }
                Read-Host "`nCredentials have been cached to $((Get-Location).Path)\ucs_cache.ucs`n`nPress any key to continue"
            }

            # Removes the UCS cache file for storing domain connection data
            4 {Remove-Item "$((Get-Location).Path)\ucs_cache.ucs"}

            # Text driven user interface for disconnecting from multiple UCS domains
            5 {
                Clear-Host
                if(!(Confirm-AnyUcsHandle)) {
                    Read-Host "There are currently no connected UCS domains`n`nPress any key to continue"
                    break
                }
                $index = 1
                $target = @{}
                Write-Host "`t`tActive Session List`n$("-"*60)"
                # Creates a hash of all active domains and prints them in a list
                foreach ($Domain in $UCS.get_keys()) {
                    if(Test-UcsHandle($UCS[$Domain])) {
                        # Composite format using -f method
                        "{0,-28} {1,20}" -f "$index) $($UCS[$Domain].Name)",$UCS[$Domain].VIP
                        $target.add($index, $UCS[$Domain].Name)
                        $index++
                    }
                }
                $command = Read-host "`nPlease Select Domains to disconnect (comma separated)"
                $disconnectList = $command.split(",")
                foreach($id in $disconnectList) {
                    # Check if entered id is within the valid range
                    if($id -lt 1 -or $id -gt $target.count) {
                        Write-Host "$id is not a valid option.  Ommitting..."
                    } else {
                    # Disconnect UCS handle
                        Write-Host "Disconnecting $($target[[int]$id])..."
                        $target[$id]
                        Disconnect-Ucs -Ucs $UCS[$target[[int]$id]].Handle
                        $script:UCS.Remove($target[[int]$id])
                    }
                }
                Read-Host "`nPress any key to return to continue"
            }

            # Disconnects all UCS domain handles
            6 {
                Clear-Host
                if(!(Confirm-AnyUcsHandle)) {
                    Read-Host "There are currently no connected UCS domains`n`nPress any key to continue"
                    break
                }
                $target = @()
                foreach    ($Domain in $UCS.get_keys()) {
                    if(Test-UcsHandle($UCS[$Domain])) {
                        Write-Host "Disconnecting $($UCS[$Domain].Name)..."
                        Disconnect-Ucs -Ucs $UCS[$Domain].Handle
                        $target += $UCS[$Domain].Name
                    }
                }
                foreach ($id in $target) {
                    $script:UCS.Remove($id)
                }
                Read-Host "`nPress any key to continue"
                break
            }

            # Returns to the main menu
            7 {return}
            default {
                Read-Host "Invalid Option.  Please select a valid option from the Menu above`nPress any key to continue"
            }
        }
    }
}

Function Get-SaveFile {
    <#
    .DESCRIPTION
        Creates a Windows File Dialog to select the save file location. Offer default filename to user.
    .PARAMETER $StartingFolder
        Default folder to present to user upon opening Windows File Save Dialog box.
    .OUTPUTS
        [string] - User selected file or silent file in script directory
    #>
    param (
        [Parameter(Mandatory)]$StartingFolder
    )
    $auto_filename = "UCS_Report_$(Get-Date -format yyyy_MM_dd_HH_mm_ss).html"
    Clear-Host
    if ($platform -match "Windows" -and !$Silent) {
        Write-Host "Opening Windows Dialog Box..."

        $null = [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms")
        $SaveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
        $SaveFileDialog.initialDirectory = $StartingFolder
        $SaveFileDialog.FileName = $auto_filename
        $SaveFileDialog.filter = "HTML Document (*.html)| *.html"
        $user_action =  $SaveFileDialog.ShowDialog()

        if ($user_action -eq 'OK') {
            $file_name = $SaveFileDialog.filename
        } else {
            Write-Host "DIALOG BOX CANCELED!! Sending output to script folder."
            $file_name = "./$($auto_filename)"
            Start-Sleep 3
        }
    } else {
        if (!$Silent) {
            Write-Host "Platform is $($platform).  Cannot use Dialog Box. Sending output to script directory."
            Start-Sleep 3
        }
        $file_name = "./$($auto_filename)"
    }
    return $file_name
}

function Get-ElapsedTime {
    <#
    .DESCRIPTION
        Computes time from start of script.
    .PARAMETER $FirstTimestamp
        Timestamp to compare against current time.
    #>
    param (
        [Parameter(Mandatory)]$FirstTimestamp
    )
    $run_time = (Get-Date) - $FirstTimestamp
    return "$($run_time.Hours)h, $($run_time.Minutes)m, $($run_time.Seconds).$(([string]$run_time.Milliseconds).PadLeft(3,'0'))s"
}

function Get-DeviceStats {
    <#
    .DESCRIPTION
        Extract specific statistics from UCS bulk dump
    .PARAMETER $UcsStats
        Collection of statistics
    .PARAMETER $DnFilter
        String for case sensitive regex match on object DN
    .PARAMETER $rnFilter
        String for case sensitive regex match on object RN
    .PARAMETER $StatList
        Array of desired returned stats (e.g., @(TotalBytes,TotalPackets))
    .PARAMETER $IsChassis
        Flag indicating stat lookup is for a chassis and that real DN must be returned. Used for NoStats operation.
    .PARAMETER $IsServer
        Flag indicating stat lookup is for a server and that real DN must be returned. Used for NoStats operation.
    #>
    param (
        [Parameter(Mandatory)] $UcsStats,
        [Parameter(Mandatory)] $DnFilter,
        [Parameter(Mandatory)] $RnFilter,
        [Parameter(Mandatory)] $StatList,
        [Switch]$IsChassis,
        [Switch]$IsServer
    )
    if ($NoStats) {
        if ($IsChassis) {
            # Get DNs for all chassis.
            $DeviceDNs = (Get-UcsChassis -Ucs $handle).Dn
        } elseif ($IsServer) {
            # Get DNs for all servers. Remove null values from resulting array, if present.
            $DeviceDNs = ((Get-UcsBlade).Dn + (Get-UcsRackUnit).Dn).Where({$null -ne $_})
        }

        # Initialize variable for set of stats
        $Data = @()

        if ($DeviceDNs) {
            # Loop through all DNs and set empty stats
            foreach ($DeviceDN in $DeviceDNs) {
                $DeviceDnData = @{}
                foreach ($Stat in $StatList) {
                    if ($Stat -eq "Dn") {$DeviceDnData[$Stat] = $DeviceDN} else {$DeviceDnData[$Stat] = 0}
                }
                $Data += $DeviceDnData
            }
        } elseif (!$IsChassis -and !$IsServer) {
            $DeviceDnData = @{}
            foreach ($Stat in $StatList) {
                $DeviceDnData[$Stat] = 0
            }
            $Data += $DeviceDnData
        }
    } else {
        # $start_time = Get-Date
        $Data = $UcsStats.Where({$_.Dn -cmatch $DnFilter -and $_.Rn -cmatch $RnFilter}) | Select-Object $StatList
        # Write-Host "Stats Lookup Time: $(Get-ElapsedTime -FirstTimestamp $start_time)"
    }

    return $Data
}

function Get-ConfiguredBootOrder {
    <#
    .DESCRIPTION
        Gets all boot policies and boot order for supplied base policy set.
    .PARAMETER $BootPolicies
        Object reference to Get-UcsBootPolicy or Get-UcsBootDefinition
    .OUTPUTS
        One or more hashtables containing data
    #>
    param (
        [Parameter(Mandatory)]$BootPolicies
    )
    # Notes regarding object nesting and where information comes from:
    #   BootPolicies (collection)
    #       |---BootPolicy (obj) - High level descriptors
    #       |   |---BootItems (collection) (unordered - needs sort)
    #       |   |   |---BootItem (obj) - Acquire high level type from Type (storage, virtual-media, lan, etc.)
    #       |   |   |   |
    #       |   |   |   |---[if]Rn=boot-security (special case) - Get SecureBoot attribute for top level block ***=== Future ===***
    #       |   |   |   |
    #       |   |   |   |---[if]virtual-media - Level1 (BootItem obj)
    #       |   |   |   |---[if]efi-shell - Level1 (BootItem obj) ***=== Future ===***
    #       |   |   |   |---[if]storage - BootItem > ItemType (child obj) > Level1 (child obj)
    #       |   |   |   |---[if]san - Level1 (BootItem obj) > Level2 (Vnics child coll) > Level3 (SAN Image child coll)
    #       |   |   |   |---[if]iscsi - Level1 (BootItem obj) > Level2 (Path child coll)
    #       |   |   |   |---[if]lan - Level1 (BootItem obj) > Level2 (Path child coll)

    $Data = @()

    # Iterate through all boot parameters for current server
    foreach ($BootPolicy in $BootPolicies) {
    # $server | Get-UcsBootDefinition | ForEach-Object {
        # Store current pipe variable to local variable
        # $BootPolicy = $_
        # Hash variable for storing current boot data
        $BootPolicyData = @{}

        # Grab informational data points from current policy
        $BootPolicyData.Dn = $BootPolicy.Dn
        $BootPolicyData.Description = $BootPolicy.Descr
        $BootPolicyData.BootMode = $BootPolicy.BootMode
        $BootPolicyData.EnforceVnicName = $BootPolicy.EnforceVnicName
        $BootPolicyData.Name = $BootPolicy.Name
        $BootPolicyData.RebootOnUpdate = $BootPolicy.RebootOnUpdate
        $BootPolicyData.Owner = $BootPolicy.Owner

        # Array variable for boot policy entries
        $BootPolicyData.Entries = @()

        # Get all child objects of the current policy and sort by boot order
        $BootItems = $BootPolicy | Get-UcsChild | Sort-Object Order
        foreach ($BootItem in $BootItems) {
            Switch ($BootItem.Type) {
                # Match local storage types
                'storage' {
                    # Get Level 1 data only from child object 2 steps down from current boot item
                    $LevelData = @{}
                    $LevelData.Level1 = $BootItem | Get-UcsChild | Get-UcsChild | Select-Object Type,Order
                    $BootPolicyData.Entries += $LevelData
                }

                # Match virtual media types
                'virtual-media' {
                    # Get Level 1 data only from current boot item
                    $LevelData = @{}
                    $LevelData.Level1 = $BootItem | Select-Object Type,Order,Access

                    # Use access level to determine media type
                    if ($LevelData.Level1.Access -match 'read-only') {
                        $LevelData.Level1.Type = 'CD/DVD'
                    } else {
                        $LevelData.Level1.Type = 'Floppy'
                    }
                    $BootPolicyData.Entries += $LevelData
                }

                # Match LAN boot types
                'lan' {
                    # Get Level 1 data from current boot item and Level 2 from child
                    $LevelData = @{}
                    $LevelData.Level1 = $BootItem | Select-Object Type,Order
                    $LevelData.Level2 = @()
                    $LevelData.Level2 += $BootItem | Get-UcsChild | Select-Object VnicName,Type
                    $BootPolicyData.Entries += $LevelData
                }
                # Match SAN boot types
                'san' {
                    # Get Level 1 data from current boot item, Level 2 from each child (Vnic), and Level 3 from
                    # children of Vnics
                    $LevelData = @{}
                    $LevelData.Level1 = $BootItem | Select-Object Type,Order
                    $LevelData.Level2 = @()

                    $Vnics = $BootItem | Get-UcsChild | Sort-Object Type
                    foreach ($Vnic in $Vnics) {
                        # Hash variable for storing current san entry data
                        $VnicData = @{}
                        $VnicData.Type = $Vnic.Type
                        $VnicData.VnicName = $Vnic.VnicName
                        $VnicData.Level3 = @()
                        $VnicData.Level3 += $Vnic | Get-UcsChild | Sort-Object Type | Select-Object Lun,Type,Wwn
                        # Add sanHash to Level2 array variable
                        $LevelData.Level2 += $VnicData
                    }
                    # Add current boot entry data to boot hash
                    $BootPolicyData.Entries += $LevelData
                }

                # Matches ISCSI boot types
                'iscsi' {
                    # Get Level 1 data from current boot item and Level 2 from child
                    $LevelData = @{}
                    $LevelData.Level1 = $BootItem | Select-Object Type,Order
                    $LevelData.Level2 = @()
                    $LevelData.Level2 += $BootItem | Get-UcsChild | Sort-Object Type | Select-Object ISCSIVnicName,Type
                    $BootPolicyData.Entries += $LevelData
                }
            }
        }
        # Sort all boot entries by Level1 Order
        $BootPolicyData.Entries = $BootPolicyData.Entries | Sort-Object {$_.Level1.Order}
        # Store boot entries to configured boot order array
        # $Data.Configured_Boot_Order += $BootPolicyData
        $Data += $BootPolicyData
    }
    return $Data
}

function Get-InventoryServerData {
    <#
    .DESCRIPTION
        Extract inventory data for blades or rackmount servers
    .PARAMETER $handle
        Handle (object reference) to target UCS domain
    .PARAMETER $AllFabricEp
        Object reference to results of Get-UcsFabricPathEp
    .PARAMETER $memoryArray
        Object reference to results of Get-UcsMemoryUnit
    .PARAMETER $EquipLocalDskDef
        Object reference to results of Get-UcsEquipmentLocalDiskDef
    .PARAMETER $EquipManufactDef
        Object reference to results of Get-UcsEquipmentManufacturingDef
    .PARAMETER $EquipPhysicalDef
        Object reference to results of Get-UcsEquipmentPhysicalkDef
    .PARAMETER $AllRunningFirmware
        Object reference to results of Get-UcsFirmwareRunning
    .PARAMETER $IsBlade
        Indicates invocation is running unique commands for blade servers. Uses rackmount search if absent.
    .OUTPUTS
        Array of hashtables
    #>

    param (
        [Parameter(Mandatory)]$handle,
        [Parameter(Mandatory)]$AllFabricEp,
        [Parameter(Mandatory)]$memoryArray,
        [Parameter(Mandatory)]$EquipLocalDskDef,
        [Parameter(Mandatory)]$EquipManufactDef,
        [Parameter(Mandatory)]$EquipPhysicalDef,
        [Parameter(Mandatory)]$AllRunningFirmware,
        [Switch]$IsBlade
    )
    $AllVirtualCircuits = Get-UcsDcxVc

    $Data = @()

    if ($IsBlade) {$servers = Get-UcsBlade -Ucs $handle} else {$servers = Get-UcsRackUnit -Ucs $handle}

    # Iterate through each server and grab relevant data
    foreach ($server in $servers) {
        # Hash variable for storing current server data
        $ServerData = @{}

        $ServerData.Dn = $Server.Dn
        $ServerData.Status = $Server.OperState
        if ($IsBlade) {
            $ServerData.Chassis = $Server.ChassisId
            $ServerData.Slot = $Server.SlotId
        } else {
            $ServerData.Rack_Id = $Server.Id
        }
        # Get Model and Description common names and format the text
        $ServerData.Model = (($EquipManufactDef.Where({$_.Sku -ieq $($Server.Model)})).Name).Replace("Cisco UCS ", "")
        $ServerData.Description = ($EquipManufactDef.Where({$_.Sku -ieq $($Server.Model)})).Description
        $ServerData.Serial = $Server.Serial
        $ServerData.Uuid = $Server.Uuid
        $ServerData.UsrLbl = $Server.UsrLbl
        $ServerData.Name = $Server.Name
        $ServerData.Service_Profile = $Server.AssignedToDn

        # If server doesn't have a service profile set profile name to Unassociated
        if(!($ServerData.Service_Profile)) {
            $ServerData.Service_Profile = "Unassociated"
        }
        # Get server child object for future iteration
        $childTargets = $Server | Get-UcsChild | Where-Object {$_.Rn -match "^bios|^mgmt|^board$"} | get-ucschild

        # Get server CPU data
        $cpu = ($childTargets | Where-Object {$_.Rn -ieq "cpu-1"}).Model
        # Get CPU common name and format text
        $ServerData.CPU_Model = "($($Server.NumOfCpus)) $($cpu -replace ('(Intel\(R\) Xeon\(R\)|CPU) ',''))"
        $ServerData.CPU_Cores = $Server.NumOfCores
        $ServerData.CPU_Threads = $Server.NumOfThreads
        # Format available memory in GB
        $ServerData.Memory = $Server.AvailableMemory/1024
        $ServerData.Memory_Speed = $Server.MemorySpeed
        # Get BIOS version. Strip uninteresting text string from right side.
        $ServerData.BIOS = ($childTargets.Where({$_.Type -eq "blade-bios"})).Version
        $ServerData.BIOS = ($ServerData.BIOS -replace ('(?!(.*\.){2}).*','')).TrimEnd('.')
        # $ServerData.BIOS = (($childTargets | Where-Object {$_.Type -eq "blade-bios"}).Version -replace ('(?!(.*\.){2}).*','')).TrimEnd('.')
        $ServerData.CIMC = ($childTargets | Where-Object {$_.Rn -ieq "fw-system"}).Version

        # =====
        # Array variable for storing server adapter data
        $ServerData.Adapters = @()

        # Iterate through each server adapter and grab detailed information
        $Adapters = $Server | Get-UcsAdaptorUnit
        foreach ($Adapter in $Adapters) {
            # Hash variable for storing current adapter data
            $AdapterData = @{}

            # Get common name of adapter and format string
            $AdapterData.Model = ($EquipManufactDef | Where-Object {$_.Sku -ieq $($adapter.Model)}).Name
            $AdapterData.Model = $AdapterData.Model -replace "Cisco UCS ", ""
            # Report adapter name field based on server type.
            if ($IsBlade) {
                $AdapterData.Name = "Adaptor-$($adapter.Id)"
            } else {
                $AdapterData.Name = "Adaptor-$($adapter.PciSlot)"
            }
            $AdapterData.Slot = $Adapter.Id
            $AdapterData.Fw = ($AllRunningFirmware.Where({$_.Deployment -eq 'system' -and $_.Dn -match $Adapter.Dn})).Version
            $AdapterData.Serial = $Adapter.Serial
            # Add current adapter hash to server adapter array
            $ServerData.Adapters += $AdapterData
        }
        # Declutter debug
        if ($Adapter) {Remove-Variable Adapter}
        if ($Adapters) {Remove-Variable Adapters}
        if ($AdapterData) {Remove-Variable AdapterData}

        # =====
        # Array variable for storing server memory data
        $ServerData.Memory_Array = @()

        # Iterage through all memory tied to current server and grab relevant data
        $Modules = $memoryArray.Where({$_.Dn -match $Server.Dn}) | Sort-Object {$_.Id -as [int]}
        foreach ($Module in $Modules) {
            # Hash variable for storing current memory data
            $ModuleData = @{}

            $ModuleData.Name = "Slot $($Module.Id)"
            $ModuleData.Location = $Module.Location
            if ($Module.Capacity -like "unspecified") {
                $ModuleData.Capacity = "empty"
                $ModuleData.Clock = "empty"
            } else {
                # Format DIMM capacity in GB
                $ModuleData.Capacity = ($Module.Capacity)/1024
                $ModuleData.Clock = $Module.Clock
            }
            $ServerData.Memory_Array += $ModuleData
        }
        # Declutter debug
        if ($Module) {Remove-Variable Module}
        if ($Modules) {Remove-Variable Modules}
        if ($ModuleData) {Remove-Variable ModuleData}

        # =====
        # Array variable for storing local storage configuration data
        $ServerData.Storage = @()

        # Iterate through each server storage controller and grab relevant data
        $Controllers = $Server | Get-UcsComputeBoard | Get-UcsStorageController
        foreach ($Controller in $Controllers) {
            # Hash variable for storing current storage controller data
            $ControllerData = @{}

            # Grab relevant controller data and store to respective controllerHash variable
            $ControllerData.ControllerStatus = $Controller.XtraProperty.ControllerStatus
            $ControllerData.Id = $Controller.Id
            $ControllerData.Model = $Controller.Model
            $ControllerData.PciAddr = $Controller.PciAddr
            $ControllerData.RaidSupport = $Controller.RaidSupport
            $ControllerData.RebuildRate = $Controller.XtraProperty.RebuildRate
            $ControllerData.Revision = $Controller.HwRevision
            $ControllerData.Serial = $Controller.Serial
            $ControllerData.Vendor = $Controller.Vendor

            # Array variable for storing controller disks
            $ControllerData.Disks = @()
            $ControllerData.Disk_Count = 0
            # Iterate through each local disk and grab relevant data
            $Disks = $Controller | Get-UcsStorageLocalDisk -Presence "equipped" | Sort-Object -Property Id
            foreach ($Disk in $Disks) {
                # Hash variable for storing current disk data
                $DiskData = @{}

                $DiskData.Blocks = $Disk.NumberOfBlocks
                $DiskData.Block_Size = $Disk.BlockSize
                $DiskData.Id = $Disk.Id
                $DiskData.Operability = $Disk.Operability
                $DiskData.Presence = $Disk.Presence
                $DiskData.Running_Version = ($AllRunningFirmware.Where({$_.Dn -match $Disk.Dn})).Version
                $DiskData.Serial = $Disk.Serial
                # Format disk size to whole GB value
                $DiskData.Size = "{0:N2}" -f ($Disk.Size/1024)
                $DiskData.Vendor = $Disk.Vendor
                $DiskData.Vid = $equipmentDef.Vid

                $DiskData.Drive_State = $Disk.XtraProperty.DiskState
                $DiskData.Link_Speed = $Disk.XtraProperty.LinkSpeed
                $DiskData.Power_State = $Disk.XtraProperty.PowerState

                # Get common name of disk model and format text
                $equipmentDef = $EquipManufactDef | Where-Object {$_.OemPartNumber -ieq $($Disk.Model)}
                $DiskData.Pid = $equipmentDef.Pid
                $DiskData.Product_Name = $equipmentDef.Name

                # Get detailed disk capability data
                $capabilities = $EquipLocalDskDef.Where({$_.Dn -match $Disk.Model})
                $DiskData.Technology = $capabilities.Technology
                $DiskData.Avg_Seek_Time = $capabilities.SeekAverageReadWrite
                $DiskData.Track_To_Seek = $capabilities.SeekTrackToTrackReadWrite

                $ControllerData.Disk_Count += 1
                # Add current disk hash to controller hash disk array
                $ControllerData.Disks += $DiskData
            }
            # Add controller hash variable to current server hash storage array
            $ServerData.Storage += $ControllerData
        }
        # Declutter debug
        if ($Controller) {Remove-Variable Controller}
        if ($Controllers) {Remove-Variable Controllers}
        if ($ControllerData) {Remove-Variable ControllerData}
        if ($Disk) {Remove-Variable Disk}
        if ($Disks) {Remove-Variable Disks}
        if ($DiskData) {Remove-Variable DiskData}

        # =====
        # Array variable for storing VIF information for current server
        $ServerData.VIFs = @()

        # Search script blocks. $_ resolved where executed
        $Search1 = {$_.Dn -Match $Server.Dn}
        $Search2 = {$_.CType -match "mux-fabric|switch-to-host"}
        $Search3 = {$_.CType -notmatch "mux-fabric(.*)?[-]"}

        # Iterate through all paths of type "mux-fabric" for the current server
        $FabricEps = $AllFabricEp | Where-Object {(& $Search1) -and (& $Search2) -and (& $Search3)}
        foreach ($FabricEp in $FabricEps) {
            # Hash variable for storing current VIF data
            $FabricEpData = @{}

            # The name of the current Path formatted to match the presentation in UCSM
            $FabricEpData.Name = "Path " + $FabricEp.SwitchId + '/' + ($FabricEp.Dn | Select-String -pattern "(?<=path[-]).*(?=[/])")[0].Matches.Value

            # Gets peer port information filtered to the current path for adapter and fex host port
            $Search2 = {$_.EpDn -match $FabricEpPeersEpDn}
            $Search3 = {$_.Dn -ne $FabricEp.Dn}
            $FabricEpPeersEpDn = ($FabricEp.EpDn | Select-String -pattern ".*(?=(.*[/]){2})").Matches.Value
            $FabricEpPeers = $AllFabricEp | Where-Object {(& $Search1) -and (& $Search2) -and (& $Search3)}

            if ($FabricEpPeers) {
                $fabric_host = $FabricEpPeers | Where-Object {$_.Rn -match "fabric.*-to-hostpc"}
                $host_adapter = $FabricEpPeers | Where-Object {$_.Rn -match "hostpc-to-adaptorpc"}

                # If Adapter PortId is greater than 1000 then format string as a port channel
                if ($host_adapter.PeerPortId -gt 1000) {
                    $FabricEpData.Adapter_Port = 'PC-' + $host_adapter.PeerPortId
                } else {
                # Else format in slot/port notation
                    $FabricEpData.Adapter_Port = "$($host_adapter.PeerSlotId)/$($host_adapter.PeerPortId)"
                }

                # If FEX PortId is greater than 1000 then format string as a port channel
                if($fabric_host.PortId -gt 1000) {
                    $FabricEpData.Fex_Host_Port = 'PC-' + $fabric_host.PortId
                } else {
                # Else format in chassis/slot/port notation
                    $FabricEpData.Fex_Host_Port = "$($fabric_host.ChassisId)/$($fabric_host.SlotId)/$($fabric_host.PortId)"
                }

                # If Network PortId is greater than 1000 then format string as a port channel
                if($FabricEp.PortId -gt 1000) {
                    $FabricEpData.Fex_Network_Port = 'PC-' + $FabricEp.PortId
                } else {
                # Else format in fabricId/slot/port notation
                    $FabricEpData.Fex_Network_Port = $FabricEp.PortId
                }

                # Server Port for current path as formatted in UCSM
                $FabricEpData.FI_Server_Port = "$($FabricEp.SwitchId)/$($FabricEp.PeerSlotId)/$($FabricEp.PeerPortId)"
            } else {
                $FabricEpData.Adapter_Port = "$($FabricEp.PeerSlotId)/$($FabricEp.PeerPortId)"
                $FabricEpData.Fex_Host_Port = "N/A"
                $FabricEpData.Fex_Network_Port = "N/A"
                $FabricEpData.FI_Server_Port = "$($FabricEp.SwitchId)/$($FabricEp.SlotId)/$($FabricEp.PortId)"
            }

            # Array variable for storing virtual circuit data
            $FabricEpData.Circuits = @()

            # Iterate through all circuits for the current vif
            $Circuits = $AllVirtualCircuits | Where-Object {$_.Dn -cmatch ($FabricEp.Dn | Select-String -pattern ".*(?<=[/])")[0].Matches.Value}
            foreach ($Circuit in $Circuits) {
                # Hash variable for storing current circuit data
                $CircuitData = @{}

                $CircuitData.Name = "Virtual Circuit $($Circuit.Id)"
                $CircuitData.vNic = $Circuit.vNic
                $CircuitData.Link_State = $Circuit.LinkState

                # Check if the current circuit is pinned to a PC uplink
                if($Circuit.OperBorderPortId -gt 0 -and $Circuit.OperBorderSlotId -eq 0) {
                    $CircuitData.FI_Uplink = "$($Circuit.SwitchId)/PC - $($Circuit.OperBorderPortId)"
                # Check if the current circuit is unpinned
                } elseif($Circuit.OperBorderPortId -eq 0 -and $Circuit.OperBorderSlotId -eq 0) {
                    $CircuitData.FI_Uplink = "unpinned"
                # Assume that the circuit is pinned to a single uplink port
                } else {
                    $CircuitData.FI_Uplink = "$($Circuit.SwitchId)/$($Circuit.OperBorderSlotId)/$($Circuit.OperBorderPortId)"
                }
                # Add current circuit data to loop array variable
                $FabricEpData.Circuits += $CircuitData
            }
            # Add vif data to server hash
            $ServerData.VIFs += $FabricEpData
        }
        # Declutter debug
        if ($FabricEp) {Remove-Variable FabricEp}
        if ($FabricEps) {Remove-Variable FabricEps}
        if ($FabricEpData) {Remove-Variable FabricEpData}
        if ($Circuit) {Remove-Variable Circuit}
        if ($Circuits) {Remove-Variable Circuits}
        if ($CircuitData) {Remove-Variable CircuitData}

        # =====
        # Get the configured boot definition of the current server
        # Array variable for storing boot order data
        $ServerData.Configured_Boot_Order = @()
        $ServerBootPolicies = $Server | Get-UcsBootDefinition
        $ServerData.Configured_Boot_Order += Get-ConfiguredBootOrder -BootPolicies $ServerBootPolicies

        # Grab actual boot order data from BIOS boot order table for current server

        # Array variable for storing boot entries
        $ServerData.Actual_Boot_Order = @()
        # Iterate through all boot entries
        $BootItems = $Server | Get-UcsBiosUnit | Get-UcsBiosBOT | Get-UcsBiosBootDevGrp | Sort-Object Order
        foreach ($BootItem in $BootItems) {
            $LevelData = @{}
            $LevelData.Descr = $BootItem.Descr

            $LevelData.Entries = @()
            $BootItem | Get-UcsBiosBootDev | ForEach-Object {$LevelData.Entries += "($($_.Order)) $($_.Descr)"}
            $ServerData.Actual_Boot_Order += $LevelData
        }
        # Add server hash data to DomainHash variable
        $Data += $ServerData
    }
    # End server Inventory Collection
    return $Data
}

function Get-SystemData {
    <#
    .DESCRIPTION
        Extract general domain information: power/temperature stats, VIP, backup policy, call home
    .PARAMETER $handle
        Handle (object reference) to target UCS domain
    .PARAMETER $DomainStatus
        Object reference to results of Get-UcsStatus
    .PARAMETER $Statistics
        Object reference to Get-UcsStatistics
    .OUTPUTS
        Hashtable
    #>

    param (
        [Parameter(Mandatory)]$handle,
        [Parameter(Mandatory)]$DomainStatus,
        [Parameter(Mandatory)]$Statistics
    )
    $Data = @{}
    $Data.Chassis_Power = @()
    $Data.Server_Power = @()
    $Data.Server_Temp = @()

    # Get UCS Cluster State
    $Data.VIP = $DomainStatus.VirtualIpv4Address
    $Data.UCSM = ($AllRunningFirmware.Where({$_.Dn -match "sys/mgmt/fw-system"})).Version
    $Data.HA_Ready = $DomainStatus.HaReady
    # Get Full State and Logical backup configuration
    $Data.Backup_Policy = (Get-UcsMgmtBackupPolicy -Ucs $handle | Select-Object AdminState).AdminState
    $Data.Config_Policy = (Get-UcsMgmtCfgExportPolicy -Ucs $handle | Select-Object AdminState).AdminState
    # Get Call Home admin state
    $Data.CallHome = (Get-UcsCallHome -Ucs $handle | Select-Object AdminState).AdminState

    # Get Chassis power statistics
    $cmd_args = @{
        UcsStats = $statistics
        DnFilter = "sys/chassis-[0-9]+/stats"
        RnFilter = "stats"
        StatList = @("Dn","InputPower","InputPowerAvg","InputPowerMax","OutputPower","OutputPowerAvg",
                     "OutputPowerMax","Suspect")
    }
    $Data.Chassis_Power += Get-DeviceStats @cmd_args -IsChassis
    $Data.Chassis_Power | ForEach-Object {$_.Dn = $_.Dn -replace ('(sys[/])|([/]stats)',"")}

    # Get Blade and Rack Mount power statistics
    $cmd_args = @{
        UcsStats = $statistics
        DnFilter = "sys/.*/board/power-stats"
        RnFilter = "power-stats"
        StatList = @("Dn","ConsumedPower","ConsumedPowerAvg","ConsumedPowerMax","InputCurrent","InputCurrentAvg",
                     "InputVoltage","InputVoltageAvg","Suspect")
    }
    $Data.Server_Power += Get-DeviceStats @cmd_args -IsServer
    $Data.Server_Power | ForEach-Object {$_.Dn = $_.Dn -replace ('(sys[/])|([/]board.*)',"")}

    # Get Blade and Rack Mount temperature statistics
    $cmd_args = @{
        UcsStats = $statistics
        DnFilter = "sys/.*/board/temp-stats"
        RnFilter = "temp-stats"
        StatList = @("Dn","FmTempSenIo","FmTempSenIoAvg","FmTempSenIoMax","FmTempSenRear","FmTempSenRearAvg",
                     "FmTempSenRearMax","FrontTemp","FrontTempAvg","FrontTempMax","Ioh1Temp","Ioh1TempAvg",
                     "Ioh1TempMax","Suspect")
    }
    $Data.Server_Temp += Get-DeviceStats @cmd_args -IsServer
    $Data.Server_Temp | ForEach-Object {$_.Dn = $_.Dn -replace ('(sys[/])|([/]board.*)',"")}

    return $Data
}
function Get-InventoryFIData {
    <#
    .DESCRIPTION
        Extract fabric interconnect inventory data
    .PARAMETER $handle
        Handle (object reference) to target UCS domain
    .PARAMETER $DomainStatus
        Object reference to results of Get-UcsStatus
    .PARAMETER $EquipManufactDef
        Object reference to results of Get-UcsEquipmentManufacturingDef
    .OUTPUTS
    #>

    param (
        [Parameter(Mandatory)]$handle,
        [Parameter(Mandatory)]$DomainStatus,
        [Parameter(Mandatory)]$EquipManufactDef
    )

    $Data = @()
    $SystemData = @{}

    $FISwitches = Get-UcsNetworkElement -Ucs $handle

    foreach ($FISwitch in $FISwitches) {

        # Hash variable for storing current FI details
        $FISwitchData = @{}

        $FISwitchData.Dn = $FISwitch.Dn
        $FISwitchData.Fabric_Id = $FISwitch.Id
        $FISwitchData.IP = $FISwitch.OobIfIp
        $FISwitchData.Operability = $FISwitch.Operability
        $FISwitchData.Thermal = $FISwitch.Thermal

        # Get leadership role and management service state
        if($FISwitch.Id -eq "A") {
            # Inventory Tab Data
            $FISwitchData.Role = $DomainStatus.FiALeadership
            $FISwitchData.State = $DomainStatus.FiAManagementServicesState
            # System Tab Data
            $SystemData.FI_A_Role = $DomainStatus.FiALeadership
            $SystemData.FI_A_IP = $FISwitch.OobIfIp
        } else {
            # Inventory Tab Data
            $FISwitchData.Role = $DomainStatus.FiBLeadership
            $FISwitchData.State = $DomainStatus.FiBManagementServicesState
            # System Tab Data
            $SystemData.FI_B_Role = $DomainStatus.FiBLeadership
            $SystemData.FI_B_IP = $FISwitch.OobIfIp
        }

        # Get the common name of the fi from the manufacturing definition and format the text
        $Model = ($EquipManufactDef | Where-Object  {$_.Sku -cmatch $($FISwitch.Model)} | Select-Object Name).Name -replace "Cisco UCS ", ""
        $FISwitchData.Model = $Model -replace "Cisco UCS ", ""

        # Array check unclear. Commenting to allow field testing without test.
        # if($fiModel -is [array]) {$FISwitchData.Model = $fiModel.Item(0) -replace "Cisco UCS ", ""} else {$FISwitchData.Model = $fiModel -replace "Cisco UCS ", ""}

        $FISwitchData.Serial = $FISwitch.Serial

        # Get FI System and Kernel FW versions
        $FIFirmware = Get-UcsFirmwareBootUnit | Where-Object {$_.Dn -match $FISwitch.Dn}
        $FISwitchData.System = ($FIFirmware | Where-Object {$_.Type -eq "system"}).Version
        $FISwitchData.Kernel = ($FIFirmware | Where-Object {$_.Type -eq "kernel"}).Version

        # Get Port licensing information
        $FILicenses = Get-UcsLicense -Ucs $handle -Scope $FISwitch.Id
        foreach ($FILicense in $FILicenses) {
            $FISwitchData.Ports_Used += $FILicense.UsedQuant
            $FISwitchData.Ports_Used += $FILicense.SubordinateUsedQuant
            $FISwitchData.Ports_Licensed += $FILicense.AbsQuant
        }

        # Get Ethernet and FC Switching mode of FI
        $FISwitchData.Ethernet_Mode = (Get-UcsLanCloud -Ucs $handle).Mode
        $FISwitchData.FC_Mode = (Get-UcsSanCloud -Ucs $handle).Mode

        # Get Local storage, VLAN, and Zone utilization numbers
        $FISwitchData.Storage = $FISwitch | Get-UcsStorageItem | Select-Object Name,Size,Used
        $Properties = @("Limit","AccessVlanPortCount","BorderVlanPortCount","AllocStatus")
        $FISwitchData.VLAN = $FISwitch | Get-UcsSwVlanPortNs | Select-Object $Properties
        $FISwitchData.Zone = $FISwitch | Get-UcsManagedObject -Classid SwFabricZoneNs | Select-Object Limit,ZoneCount,AllocStatus

        # Sort Expression to filter port id to be just the numerical port number and sort ascending
        $SortExpression = {if ($_.Dn -match "(?=port[-]).*") {($matches[0] -replace ".*(?<=[-])",'') -as [int]}}
        # Get Fabric Port Configuration and sort by port id using the above sort expression
        $Properties = @("AdminState","Dn","IfRole","IfType","LicState","LicGP","Mac","Mode","OperState","OperSpeed","XcvrType","PeerDn","PeerPortId","PeerSlotId","PortId","SlotId","SwitchId")
        $FISwitchData.Ports = Get-UcsFabricPort -Ucs $handle -SwitchId $FISwitch.Id -AdminState enabled | Sort-Object $SortExpression | Select-Object $Properties
        $Properties = @("AdminState","Dn","IfRole","IfType","LicState","LicGP","Wwn","Mode","OperState","OperSpeed","XcvrType","PortId","SlotId","SwitchId")
        $FISwitchData.FcUplinkPorts = Get-UcsFiFcPort -Ucs $handle -SwitchId "$($FISwitch.Id)" -AdminState 'enabled' | Sort-Object $SortExpression | Select-Object $Properties

        # Store fi hash to domain hash variable
        $Data += $FISwitchData

    }
    return $Data,$SystemData
}

function Get-InventoryChassisData {
    <#
    .DESCRIPTION
        Extract inventory data for chassis including basic member blade data
    .PARAMETER $handle
        Handle (object reference) to target UCS domain
    .PARAMETER $EquipPhysicalDef
        Object reference to results of Get-UcsEquipmentPhysicalkDef
    .OUTPUTS
    #>

    param (
        [Parameter(Mandatory)]$handle,
        [Parameter(Mandatory)]$EquipPhysicalDef
    )

    $Data = @()

    # Iterate through chassis inventory and grab relevant data
    $AllChassis = Get-UcsChassis -Ucs $handle
    foreach ($Chassis in $AllChassis) {
        # Hash variable for storing current chassis data
        $ChassisData = @{}

        $ChassisData.Dn = $Chassis.Dn
        $ChassisData.Id = $Chassis.Id
        $ChassisData.Model = $Chassis.Model
        $ChassisData.Status = $Chassis.OperState
        $ChassisData.Operability = $Chassis.Operability
        $ChassisData.Power = $Chassis.Power
        $ChassisData.Thermal = $Chassis.Thermal
        $ChassisData.Serial = $Chassis.Serial

        $ChassisData.Blades = @()

        # Initialize chassis used slot count to 0
        $UsedSlots = 0
        # Iterate through all blades within current chassis
        $Blades = Get-UcsBlade -ChassisId $Chassis.Id
        foreach ($Blade in $Blades) {
            # Hash variable for storing current blade data
            $BladeData = @{}

            $BladeData.Model = $Blade.Model
            $BladeData.SlotId = $Blade.SlotId
            $BladeData.Service_Profile = $Blade.AssignedToDn
            # Get width of blade and convert to slot count
            $BladeData.Width = [math]::floor((($EquipPhysicalDef | Where-Object {$_.Dn -ilike "*$($Blade.Model)*"}).Width)/8)
            # Increment used slot count by current blade width
            $UsedSlots += $BladeData.Width
            $ChassisData.Blades += $BladeData
        }
        # Get Used slots and slots available from iterated slot count
        $ChassisData.SlotsUsed = $UsedSlots
        $ChassisData.SlotsAvailable = 8 - $UsedSlots

        # Get chassis PSU data and redundancy mode
        $ChassisData.Psus = @()
        $ChassisData.Psus = $Chassis | Get-UcsPsu | Sort-Object Id | Select-Object Type,Id,Model,Serial,Dn
        $ChassisData.Power_Redundancy = ($Chassis | Get-UcsComputePsuControl | Select-Object Redundancy).Redundancy

        # Add chassis to domain hash variable
        $Data += $ChassisData
    }
    return $Data
}

function Get-InventoryIOModuleData {
    <#
    .DESCRIPTION
        Extract inventory data for I/O Modules (FEX)
    .PARAMETER $handle
        Handle (object reference) to target UCS domain
    .PARAMETER $EquipManufactDef
        Object reference to results of Get-UcsEquipmentManufacturingDef
    .PARAMETER $AllRunningFirmware
        Object reference to results of Get-UcsFirmwareRunning
    .OUTPUTS
    #>

    param (
        [Parameter(Mandatory)]$handle,
        [Parameter(Mandatory)]$EquipManufactDef,
        [Parameter(Mandatory)]$AllRunningFirmware
    )

    $Data = @()

    # Get all fabric and blackplane ports for future iteration
    $FabricPorts = Get-UcsEtherSwitchIntFIo -Ucs $handle
    $BackplanePorts = Get-UcsEtherServerIntFIo -Ucs $handle

    # Iterate through each IOM and grab relevant data
    $Ioms = Get-UcsIom -Ucs $handle
    foreach ($Iom in $Ioms) {
        # Hash variable for storing current IOM data
        $IomData = @{}

        $IomData.Dn = $Iom.Dn
        $IomData.Chassis = $Iom.ChassisId
        $IomData.Fabric_Id = $Iom.SwitchId
        $IomData.Serial = $Iom.Serial

        # Get common name of IOM model and format for viewing
        $IomData.Model = ($EquipManufactDef | Where-Object {$_.Sku -cmatch $($Iom.Model)}).Name -replace "Cisco UCS ", ""

        # Get the IOM uplink port channel name if configured
        $IomData.Channel = (Get-UcsPortGroup -Ucs $handle -Dn "$($Iom.Dn)/fabric-pc" | Get-UcsEtherSwitchIntFIoPc).Rn

        # Get IOM running and backup fw versions
        $IomData.Running_FW = ($AllRunningFirmware.Where({$_.Dn -match "sys/mgmt/fw-system"})).Version
        $IomData.Backup_FW = (Get-UcsFirmwareUpdatable -Filter "Dn -cmatch $($Iom.Dn)").Version

        # IOM Port Filter
        $ObjectFilter = {$_.ChassisId -eq "$($Iom.ChassisId)" -and $_.SwitchId -eq "$($Iom.SwitchId)"}

        # Initialize FabricPorts array for storing IOM port data
        $IomData.FabricPorts = @()

        # Iterate through all fabric ports tied to the current IOM
        $IomFabricPorts = $FabricPorts | Where-Object $ObjectFilter | Sort-Object -Property PortId
        foreach ($IomFabricPort in $IomFabricPorts) {
            # Hash variable for storing current fabric port data
            $PortData = @{}

            $PortData.Name = 'Fabric Port ' + $IomFabricPort.SlotId + '/' + $IomFabricPort.PortId
            $PortData.OperState = $IomFabricPort.OperState
            $PortData.PortChannel = $IomFabricPort.EpDn
            $PortData.PeerSlotId = $IomFabricPort.PeerSlotId
            $PortData.PeerPortId = $IomFabricPort.PeerPortId
            $PortData.FabricId = $IomFabricPort.SwitchId
            $PortData.Ack = $IomFabricPort.Ack
            $PortData.Peer = $IomFabricPort.PeerDn
            # Add current fabric port hash variable to FabricPorts array
            $IomData.FabricPorts += $PortData
        }
        # Initialize BackplanePorts array for storing IOM port data
        $IomData.BackplanePorts = @()

        # Iterate through all backplane ports tied to the current IOM
        $IomBackplanePorts = $BackplanePorts | Where-Object $ObjectFilter | Sort-Object -Property PortId
        foreach ($IomBackplanePort in $IomBackplanePorts) {
            # Hash variable for storing current backplane port data
            $PortData = @{}

            $PortData.Name = 'Backplane Port ' + $IomBackplanePort.SlotId + '/' + $IomBackplanePort.PortId
            $PortData.OperState = $IomBackplanePort.OperState
            $PortData.PortChannel = $IomBackplanePort.EpDn
            $PortData.FabricId = $IomBackplanePort.SwitchId
            $PortData.Peer = $IomBackplanePort.PeerDn
            # Add current backplane port hash variable to FabricPorts array
            $IomData.BackplanePorts += $PortData
        }
        # Add IOM to domain hash variable
        $Data += $IomData
    }
    return $Data
}

function Get-PolicyData {
    <#
    .DESCRIPTION
        Extract policy data
    .PARAMETER $handle
        Handle (object reference) to target UCS domain
    .PARAMETER $MaintenancePolicies
        Object reference to results of Get-UcsMaintenancePolicy
    .OUTPUTS
    #>

    param (
        [Parameter(Mandatory)]$handle,
        [Parameter(Mandatory)]$MaintenancePolicies
    )

    $Data = @{}
    $Data.SystemPolicies = @{}
    $Data.Mgmt_IP_Pool = @{}

    # Grab DNS and NTP data
    $Data.SystemPolicies.DNS = @()
    $Data.SystemPolicies.DNS += (Get-UcsDnsServer -Ucs $handle).Name | Select-Object -Unique
    $Data.SystemPolicies.NTP = @()
    $Data.SystemPolicies.NTP += (Get-UcsNtpServer -Ucs $handle).Name | Select-Object -Unique
    $Data.SystemPolicies.Timezone = (Get-UcsTimezone -Ucs $handle).Timezone | Select-Object -Unique

    # Get chassis discovery data
    $Chassis_Discovery = Get-UcsChassisDiscoveryPolicy -Ucs $handle
    $Data.SystemPolicies.Action = $Chassis_Discovery.Action
    $Data.SystemPolicies.Grouping = $Chassis_Discovery.LinkAggregationPref

    $Data.SystemPolicies.Power = (Get-UcsPowerControlPolicy -Ucs $handle).Redundancy
    $Data.SystemPolicies.FirmwareAutoSyncAction = (Get-UcsFirmwareAutoSyncPolicy -Ucs $handle).SyncState
    $Data.SystemPolicies.Maint = ($MaintenancePolicies.Where({$_.Name -eq 'default'})).UptimeDisr

    # Maintenance Policies
    $Data.Maintenance = @()
    $Data.Maintenance += $MaintenancePolicies | Select-Object Name,Dn,UptimeDisr,Descr,SchedName

    # Host Firmware Packages
    $Data.FW_Packages = @()
    $Data.FW_Packages += Get-UcsFirmwareComputeHostPack -Ucs $handle | Select-Object Name,BladeBundleVersion,RackBundleVersion

    # LDAP Policy Data
    $Data.LDAP_Providers = @()
    $Data.LDAP_Providers += Get-UcsLdapProvider -Ucs $handle | Select-Object Name,Rootdn,Basedn,Attribute
    $Data.LDAP_Mappings = @()
    $LdapGroupMaps = Get-UcsLdapGroupMap -Ucs $handle
    foreach ($LdapGroupMap in $LdapGroupMaps) {
        $LdapGroupMapData = @{}
        $LdapGroupMapData.Name = $LdapGroupMap.Name
        $LdapGroupMapData.Roles = ($LdapGroupMap | Get-UcsUserRole).Name
        $LdapGroupMapData.Locales = ($LdapGroupMap | Get-UcsUserLocale).Name
        $Data.LDAP_Mappings += $LdapGroupMapData
    }

    # Boot Order Policies
    $Data.Boot_Policies = @()
    $AllBootPolicies = Get-UcsBootPolicy -Ucs $handle
    $Data.Boot_Policies += Get-ConfiguredBootOrder -BootPolicies $AllBootPolicies

    # Get the default external management pool
    $ExtMgmtPoolBlock = Get-UcsIpPoolBlock -Ucs $handle -Filter "Dn -cmatch ext-mgmt"
    $Data.Mgmt_IP_Pool.From = $ExtMgmtPoolBlock.From
    $Data.Mgmt_IP_Pool.To = $ExtMgmtPoolBlock.To

    $ExtMgmtPoolParent = $ExtMgmtPoolBlock | Get-UcsParent
    $Data.Mgmt_IP_Pool.Size = $ExtMgmtPoolParent.Size
    $Data.Mgmt_IP_Pool.Assigned = $ExtMgmtPoolParent.Assigned

    # Mgmt IP Allocation
    $Data.Mgmt_IP_Allocation = @()
    $Assignments = $ExtMgmtPoolParent | Get-UcsIpPoolPooled -Filter "Assigned -ieq yes"
    foreach ($Assignment in $Assignments) {
        $AssignmentData = @{}
        $AssignmentData.Dn = $Assignment.AssignedToDn -replace "/mgmt/*.*", ""
        $AssignmentData.IP = $Assignment.Id
        $AssignmentData.Subnet = $Assignment.Subnet
        $AssignmentData.GW = $Assignment.DefGw
        $Data.Mgmt_IP_Allocation += $AssignmentData
    }

    # UUID
    $Data.UUID_Pools = @()
    $Data.UUID_Pools += Get-UcsUuidSuffixPool -Ucs $handle | Select-Object Dn,Name,AssignmentOrder,Prefix,Size,Assigned
    $Data.UUID_Assignments = @()
    $Data.UUID_Assignments += Get-UcsUuidpoolAddr -Ucs $handle -Assigned yes | select-object AssignedToDn,Id | sort-object -property AssignedToDn

    # Server Pools
    $Data.Server_Pools = @()
    $Data.Server_Pools += Get-UcsServerPool -Ucs $handle | Select-Object Dn,Name,Size,Assigned
    $Data.Server_Pool_Assignments = @()
    $Data.Server_Pool_Assignments += Get-UcsServerPoolAssignment -Ucs $handle | Select-Object Name,AssignedToDn

    return $Data
}

function Get-ServiceProfileGeneralData {
    param (
        $ServiceProfile,
        [Parameter(Mandatory)]$MaintenancePolicies,
        [Parameter(Mandatory)]$SPTemplateType,
        [Parameter(Mandatory)]$SPTemplateName
    )

    $Data = @{}

    # Workaround solution. Monolithic function blah
    if (!$ServiceProfile) {$ServiceProfile = [System.Management.Automation.Internal.AutomationNull]::Value}

    if ($ServiceProfile.Type -match "instance") {
        # Regular Service Profiles
        $Data.Name = $ServiceProfile.Name
        $Data.Overall_Status = $ServiceProfile.operState
        $Data.AssignState = $ServiceProfile.AssignState
        $Data.AssocState = $ServiceProfile.AssocState
        $Data.Power_State = ($ServiceProfile | Get-UcsChild -ClassId LsPower | Select-Object State).State
        $Data.UserLabel = $ServiceProfile.UsrLbl
        $Data.Descr = $ServiceProfile.Descr
        $Data.Owner = $ServiceProfile.PolicyOwner
        $Data.Uuid = $ServiceProfile.Uuid
        $Data.UuidPool = $ServiceProfile.OperIdentPoolName
        $Data.Associated_Server = $ServiceProfile.PnDn
        $Data.Template_Name = $SPTemplateName
        $Data.Template_Instance = $ServiceProfile.OperSrcTemplName

        $pool = $ServiceProfile | Get-UcsServerPoolAssignment
        $Data.Assignment = @{}
        if($pool.Count -gt 0) {
            $Data.Assignment.Server_Pool = $pool.Name
            $Data.Assignment.Qualifier = $pool.Qualifier
            $Data.Assignment.Restrict_Migration = $pool.RestrictMigration
        } else {
            $lsServer = $ServiceProfile | Get-UcsLsBinding
            $Data.Assignment.Server = $lsServer.AssignedToDn
            $Data.Assignment.Restrict_Migration = $lsServer.RestrictMigration
        }

    } else {
        # Service Profile Templates
        $Data.Name = $SPTemplateName
        $Data.Type = $SPTemplateType
        $Data.Description = $ServiceProfile.Descr
        $Data.UUIDPool = $ServiceProfile.IdentPoolName
        $Data.Boot_Policy = $ServiceProfile.OperBootPolicyName
        $Data.PowerState = ($ServiceProfile | Get-UcsServerPower).State
        $Data.MgmtAccessPolicy = $ServiceProfile.ExtIPState
        $Data.Server_Pool = $ServiceProfile | Get-UcsServerPoolAssignment | Select-Object Name,Qualifier,RestrictMigration

        if ($ServiceProfile.Type -match "(updating|initial)-template") {
            $Data.Maintenance_Policy = $MaintenancePolicies.Where({$_.Dn -eq $ServiceProfile.OperMaintPolicyName}) | Select-Object Name,Dn,Descr,UptimeDisr
        } else {
        $Data.Maintenance_Policy = ""
        }
    }

    return $Data
}

function Get-ServiceProfileStorageData {
    param (
        $ServiceProfile,
        [Parameter(Mandatory)]$VnicFcNode,
        [Parameter(Mandatory)]$VnicConnDef
    )

    $Data = @{}

    # Workaround solution. Monolithic function blah
    if (!$ServiceProfile) {$ServiceProfile = [System.Management.Automation.Internal.AutomationNull]::Value}

    if ($ServiceProfile) {
        $Data.Nwwn = ($VnicFcNode.Where({$_.Dn -match $ServiceProfile.Dn})).Addr
        $Data.Nwwn_Pool = ($VnicFcNode.Where({$_.Dn -match $ServiceProfile.Dn})).IdentPoolName
        $Properties = @("Name","Mode","ProtectConfig","XtraProperty","FlexFlashRAIDReportingState","FlexFlashState")
        $Data.Local_Disk_Config = Get-UcsLocalDiskConfigPolicy -Dn $ServiceProfile.OperLocalDiskPolicyName | Select-Object $Properties
        $Data.Connectivity_Policy = ($VnicConnDef.Where({$_.Dn -match $ServiceProfile.Dn})).SanConnPolicyName
        $Data.Connectivity_Instance = ($VnicConnDef.Where({$_.Dn -match $ServiceProfile.Dn})).OperSanConnPolicyName
    } else {
        $Data.Nwwn = $null
        $Data.Nwwn_Pool = $null
        $Data.Connectivity_Policy = $null
        $Data.Connectivity_Instance = $null
        $Data.Local_Disk_Config = ""
    }

    # Array variable for storing HBA data
    $Data.Hbas = @()

    $Hbas = $ServiceProfile | Get-UcsVhba -Ucs $handle
    foreach ($Hba in $Hbas) {
        $HbaData = @{}

        $HbaData.Name = $Hba.Name
        $HbaData.FabricId = $Hba.SwitchId
        $HbaData.Desired_Order = $Hba.Order
        $HbaData.Actual_Order = $Hba.OperOrder
        $HbaData.Desired_Placement = $Hba.AdminVcon
        $HbaData.Actual_Placement = $Hba.OperVcon

        if ($ServiceProfile.Type -eq 'instance') {
            $HbaData.Pwwn = $Hba.Addr
            $HbaData.EquipmentDn = $Hba.EquipmentDn
            $HbaData.Vsan = ($Hba | Get-UcsChild).OperVnetName
        } else {
            $HbaData.Pwwn = $Hba.IdentPoolName
            $HbaData.Vsan = ($Hba | Get-UcsChild).Name
        }
        $Data.Hbas += $HbaData
    }

    return $Data
}

function Get-ServiceProfileNetworkData {
    param (
        $ServiceProfile,
        [Parameter(Mandatory)]$VnicConnDef
    )

    $Data = @{}

    # Workaround solution. Monolithic function blah
    if (!$ServiceProfile) {$ServiceProfile = [System.Management.Automation.Internal.AutomationNull]::Value}

    # Lan Connectivity Policy
    if ($ServiceProfile) {
        $Data.Connectivity_Policy = ($VnicConnDef.Where({$_.Dn -match $ServiceProfile.Dn})).LanConnPolicyName
    } else {
        $Data.Connectivity_Policy = $null
    }

    if ($ServiceProfile.Type -ne 'instance') {$Data.DynamicVnic_Policy = $ServiceProfile.DynamicConPolicyName}

    # Array variable for storing
    $Data.Nics = @()

    # Iterate through each NIC and grab configuration details
    $Nics = $ServiceProfile | Get-UcsVnic
    foreach ($Nic in $Nics) {
        $NicData = @{}

        $NicData.Name = $Nic.Name
        $NicData.Mac_Address = $Nic.Addr
        $NicData.Desired_Order = $Nic.Order
        $NicData.Actual_Order = $Nic.OperOrder
        $NicData.Fabric_Id = $Nic.SwitchId
        $NicData.Desired_Placement = $Nic.AdminVcon
        $NicData.Actual_Placement = $Nic.OperVcon
        $NicData.Adaptor_Profile = $Nic.AdaptorProfileName
        $NicData.Control_Policy = $Nic.NwCtrlPolicyName
        if ($ServiceProfile.Type -eq 'instance') {
            $NicData.Qos = $Nic.OperQosPolicyName
            $NicData.Mtu = $Nic.Mtu
            $NicData.EquipmentDn = $Nic.EquipmentDn
        }

        # Array for storing VLANs
        $NicData.Vlans = @()

        # Grab all VLANs
        # TODO: Call yields nothing when using VLAN groups. Create group drill down. Group class ID is FabricNetGroupRef
        $NicData.Vlans += $Nic | Get-UcsChild -ClassId VnicEtherIf | Select-Object OperVnetName,Vnet,DefaultNet | Sort-Object {($_.Vnet) -as [int]}
            # Random thoughts
            # # Get VLAN from vNIC (standard)
            # Get-UcsVnic -Dn org-root/ls-SPT-B200M3/ether-vminc4 | Get-UcsChild | ForEach-Object {Get-UcsVlan -Dn $_.OperVnetDn | Select-Object Dn}

            # Get-UcsVnic -Dn <vnic path> | Get-UcsChild | ForEach-Object {Get-UcsFabricNetGroup -Dn $_.OperName}
            # Get-UcsFabricPooledVlan -Filter "Dn -cmatch <fab net grp path>" | foreach <get-ucsvlan blah>
        $Data.Nics += $NicData
    }

    return $Data
}

function Get-ServiceProfileIscsiData {
    param (
        $ServiceProfile
    )

    $Data = @()

    # Workaround solution. Monolithic function blah
    if (!$ServiceProfile) {$ServiceProfile = [System.Management.Automation.Internal.AutomationNull]::Value}

    # Iterate through iSCSI interface configuration
    $Nics = $ServiceProfile | Get-UcsVnicIscsi
    foreach ($Nic in $Nics) {
        $NicData = @{}

        $NicData.Name = $Nic.Name
        $NicData.Overlay = $Nic.VnicName
        $NicData.Iqn = $Nic.InitiatorName
        $NicData.Adapter_Policy = $Nic.AdaptorProfileName
        $NicData.Mac = $Nic.Addr
        $NicData.Vlan = ($Nic | Get-UcsVnicVlan).VlanName

        $Data += $NicData
    }

    return $Data
}

function Get-ServiceProfileBootOrderData {
}

function Get-ServiceProfilePoliciesData {
    param (
        $ServiceProfile
    )

    $Data = @{}
    $Data.Bios = $ServiceProfile.BiosProfileName
    $Data.Fw = $ServiceProfile.HostFwPolicyName
    $Data.Ipmi = $ServiceProfile.MgmtAccessPolicyName
    $Data.Power = $ServiceProfile.PowerPolicyName
    $Data.Scrub = $ServiceProfile.ScrubPolicyName
    $Data.Sol = $ServiceProfile.SolPolicyName
    $Data.Stats = $ServiceProfile.StatsPolicyName

    return $Data
}

function Get-ServiceProfileData {
    <#
    .DESCRIPTION

    .PARAMETER $handle
        Handle (object reference) to target UCS domain
    .PARAMETER $Statistics
        Object reference to Get-UcsStatistics
    .OUTPUTS
    #>

    param (
        [Parameter(Mandatory)]$handle,
        [Parameter(Mandatory)]$Statistics,
        [Parameter(Mandatory)]$MaintenancePolicies
    )

    $Data = @{}

    # Get data from UCS domain once instead of repeated calls within loops
    # Service Profile Collection - Includes both Templates and Instances
    $AllSPs = Get-UcsServiceProfile -Ucs $handle
    # Node WWN Collection - Query for WWNN assignment to Service Profile
    $VnicFcNode = Get-UcsVnicFcNode -Ucs $handle
    # LAN Connectivity Policy Collection - Query for Policy assignment to Service Profile
    $VnicConnDef = Get-UcsVnicConnDef -Ucs $handle

    # Create array of Service Profile Template DNs + 1 blank. Blank is used for Service Profiles with no template.
    $SPTemplateDNs = @()
    $SPTemplateDNs += ($AllSPs.Where({$_.Type -match "(updating|initial)-template"})).Dn
    $SPTemplateDNs += ""

    # Iterate DN array and acquire configuration data
    foreach ($SPTemplateDN in $SPTemplateDNs) {
        # Get service profile object by DN string
        # Blank DN yields [System.Management.Automation.Internal.AutomationNull]::Value
        $SPTemplate = $AllSPs.Where({$_.Dn -eq "$SPTemplateDN"})

        $SPTemplateName = if ($SPTemplate) {$SPTemplate.Name} else {"Unbound"}

        # Temp var. Populates $domainHash.Profiles.<SPT Dn>
        $SPTemplateData = @{}

        # Switch statement to format the template type
        switch ($SPTemplate.Type) {
            "updating-template" {$SPTemplateType = "Updating"}
            "initial-template" {$SPTemplateType = "Initial"}
            default {$SPTemplateType = "N/A"}
        }

        # TODO: This data is duplicated: top level and under General
        $SPTemplateData.Type = $SPTemplateType

        # Template Details - General Tab
        # Temp var. Populates $domainHash.Profiles.<SPT Dn>.General
        $SPTemplateData.General = @{}
        $cmd_args = @{
            ServiceProfile = $SPTemplate
            MaintenancePolicies = $MaintenancePolicies
            SPTemplateType = $SPTemplateType
            SPTemplateName = $SPTemplateName
        }
        $SPTemplateData.General = Get-ServiceProfileGeneralData @cmd_args

        # Template Details - Policies Tab
        # Temp var. Populates $domainHash.Profiles.<SPT Dn>.Policies
        $SPTemplateData.Policies = @{}
        $SPTemplateData.Policies = Get-ServiceProfilePoliciesData -ServiceProfile $SPTemplate

        # Template Details - Storage Tab
        # Temp var. Populates $domainHash.Profiles.<SPT Dn>.Storage
        $SPTemplateData.Storage = @{}
        $cmd_args = @{
            ServiceProfile = $SPTemplate
            VnicFcNode = $VnicFcNode
            VnicConnDef = $VnicConnDef
        }
        $SPTemplateData.Storage = Get-ServiceProfileStorageData @cmd_args

        # Template Details - Network Tab
        # Temp var. Populates $domainHash.Profiles.<SPT Dn>.Network
        $SPTemplateData.Network = @{}
        $cmd_args = @{
            ServiceProfile = $SPTemplate
            VnicConnDef = $VnicConnDef
        }
        $SPTemplateData.Network = Get-ServiceProfileNetworkData @cmd_args

        # Template Details - iSCSI vNICs Tab
        # Temp var. Populates $domainHash.Profiles.<SPT Dn>.iSCSI
        $SPTemplateData.iSCSI = @()
        $SPTemplateData.iSCSI += Get-ServiceProfileIscsiData -ServiceProfile $SPTemplate

        # Service Profile Instances (Only instances bound to current template.)
        # Temp var. Populates $domainHash.Profiles.<SPT Dn>.Profiles
        $SPTemplateData.Profiles = @()

        # Iterate collection of bound service profiles
        $BoundSPs = $AllSPs.Where({$_.OperSrcTemplName -ieq "$SPTemplateDN" -and $_.Type -ieq "instance"})
        foreach ($BoundSP in $BoundSPs) {
            # Temp var. Populates $domainHash.Profiles.<SPT Dn>.Profiles.<SP Array>.<item>
            $BoundSPData = @{}

            # Service Profiles Tab Tables
            $BoundSPData.Dn = $BoundSP.Dn
            $BoundSPData.Service_Profile = $BoundSP.Name
            $BoundSPData.UsrLbl = $BoundSP.UsrLbl
            $BoundSPData.Assigned_Server = $BoundSP.PnDn
            $BoundSPData.Assoc_State = $BoundSP.AssocState
            $BoundSPData.Maint_Policy = $BoundSP.MaintPolicyName
            $BoundSPData.Maint_PolicyInstance = $BoundSP.OperMaintPolicyName
            $BoundSPData.FW_Policy = $BoundSP.HostFwPolicyName
            $BoundSPData.BIOS_Policy = $BoundSP.BiosProfileName
            $BoundSPData.Boot_Policy = $BoundSP.OperBootPolicyName

            # Service Profile Details Modal - General Tab
            # Temp var. Populates $domainHash.Profiles.<SPT Dn>.Profiles.<SP Array>.<item>.General
            $BoundSPData.General = @{}
            $cmd_args = @{
                ServiceProfile = $BoundSP
                MaintenancePolicies = $MaintenancePolicies
                SPTemplateType = $SPTemplateType
                SPTemplateName = $SPTemplateName
            }
            $BoundSPData.General = Get-ServiceProfileGeneralData @cmd_args

            # Service Profile Details Modal - Policies Tab
            # Temp var. Populates $domainHash.Profiles.<SPT Dn>.Profiles.<SP Array>.<item>.Policies
            $BoundSPData.Policies = @{}
            $BoundSPData.Policies = Get-ServiceProfilePoliciesData -ServiceProfile $BoundSP

            # Service Profile Details Modal - Storage Tab
            # Temp var. Populates $domainHash.Profiles.<SPT Dn>.Profiles.<SP Array>.<item>.Storage
            $BoundSPData.Storage = @{}
            $cmd_args = @{
                ServiceProfile = $BoundSP
                VnicFcNode = $VnicFcNode
                VnicConnDef = $VnicConnDef
            }
            $BoundSPData.Storage = Get-ServiceProfileStorageData @cmd_args

            # Service Profile Details Modal - Network Tab
            # Temp var. Populates $domainHash.Profiles.<SPT Dn>.Profiles.<SP Array>.<item>.Network
            $BoundSPData.Network = @{}
            $cmd_args = @{
                ServiceProfile = $BoundSP
                VnicConnDef = $VnicConnDef
            }
            $BoundSPData.Network = Get-ServiceProfileNetworkData @cmd_args

            # Service Profile Details Modal - iSCSI Tab
            # Temp var. Populates $domainHash.Profiles.<SPT Dn>.Profiles.<SP Array>.<item>.iSCSI
            $BoundSPData.iSCSI = @()
            $BoundSPData.iSCSI += Get-ServiceProfileIscsiData -ServiceProfile $BoundSP

            # Service Profile Details Modal - Performance Tab
            # Temp var. Populates $domainHash.Profiles.<SPT Dn>.Profiles.<SP Array>.<item>.Performance
            $BoundSPData.Performance = @{}

            # Acquire data for associated Service Profiles only (i.e., bound to hardware)
            if($BoundSPData.Assoc_State -eq 'associated') {
                # TODO: Migrate this data to top level ($domainHash.Collection). Doesn't make sense here.
                # Get the collection time interval for adapter performance
                $interval = (Get-UcsCollectionPolicy -Name "adapter" | Select-Object CollectionInterval).CollectionInterval
                # Normalize collection interval to seconds
                Switch -wildcard (($interval -split '[0-9]')[-1]) {
                    "minute*" {$BoundSPData.Performance.Interval = ((($interval -split '[a-z]')[0]) -as [int]) * 60}
                    "second*" {$BoundSPData.Performance.Interval = ((($interval -split '[a-z]')[0]) -as [int])}
                }

                $cmd_args = @{
                    UcsStats = $Statistics
                    RnFilter = "vnic-stats"
                    StatList = @("BytesRx","BytesRxDeltaAvg","BytesTx","BytesTxDeltaAvg","PacketsRx","PacketsRxDeltaAvg","PacketsTx","PacketsTxDeltaAvg")
                }
                # Iterate through each vHBA and grab performance data
                $BoundSPData.Performance.vHbas = @{}
                $BoundSPData.Storage.Hbas | ForEach-Object {
                    $BoundSPData.Performance.vHbas[$_.Name] = Get-DeviceStats @cmd_args -DnFilter $_.EquipmentDn
                }
                # Iterate through each vNIC and grab performance data
                $BoundSPData.Performance.vNics = @{}
                $BoundSPData.Network.Nics | ForEach-Object {
                    $BoundSPData.Performance.vNics[$_.Name] = Get-DeviceStats @cmd_args -DnFilter $_.EquipmentDn
                }
            }

            # Add current data to array
            $SPTemplateData.Profiles += $BoundSPData
        }

        # Add current data to hashtable
        $Data[$SPTemplateDN] = $SPTemplateData
    }
    return $Data
}

function Get-LanTabData {
    <#
    .DESCRIPTION
        Extract SAN tab data: FC/FCOE uplinks and FC/FCOE/NAS storage ports
    .PARAMETER $handle
        Handle (object reference) to target UCS domain
    .PARAMETER $UcsFIs
        List of hashtables containing previously collected FI inventory data (Get-InventoryFIData)
    .PARAMETER $Statistics
        Object reference to Get-UcsStatistics
    .OUTPUTS
    #>

    param (
        [Parameter(Mandatory)]$handle,
        [Parameter(Mandatory)]$UcsFIs,
        [Parameter(Mandatory)]$Statistics
    )

    $Data = @{}
    $Data.UplinkPorts = @()
    $Data.ServerPorts = @()

    # Iterate through each FI and collect port performance data based on port role
    foreach ($UcsFI in $UcsFIs) {
        # Uplink and Server Ports
        $Ports = $UcsFI.Ports
        foreach ($Port in $Ports) {
            $PortData = @{}
            $PortData.Dn = $Port.Dn
            $PortData.PortId = $Port.PortId
            $PortData.SlotId = $Port.SlotId
            $PortData.Fabric_Id = $Port.SwitchId
            $PortData.Mac = $Port.Mac
            $PortData.Speed = $Port.OperSpeed
            $PortData.IfType = $Port.IfType
            $PortData.IfRole = $Port.IfRole
            $PortData.XcvrType = $Port.XcvrType
            $PortData.Performance = @{}
            $cmd_args = @{
                UcsStats = $Statistics
                DnFilter = "$($port.Dn)/.*stats"
                StatList = @("TotalBytes","TotalPackets","TotalBytesDeltaAvg")
            }
            $PortData.Performance.Rx = Get-DeviceStats @cmd_args -RnFilter "rx[-]stats"
            $PortData.Performance.Tx = Get-DeviceStats @cmd_args -RnFilter "tx[-]stats"
            $PortData.Status = $Port.OperState
            $PortData.State = $Port.AdminState

            # Store uplinks separately from server ports
            if ($Port.IfRole -cmatch "network") {
                $Data.UplinkPorts += $PortData
            } elseif ($Port.IfRole -cmatch "server") {
                $Data.ServerPorts += $PortData
            }
        }
    }

    # Fabric PortChannels
    $Data.FabricPcs = @()
    $PortChannels = Get-UcsFabricServerPortChannel -Ucs $handle
    foreach ($PortChannel in $PortChannels) {
    # Get-UcsFabricServerPortChannel -Ucs $handle | ForEach-Object {
        $PortChannelData = @{}
        $PortChannelData.Name = $PortChannel.Rn
        $PortChannelData.Chassis = $PortChannel.ChassisId
        $PortChannelData.Fabric_Id = $PortChannel.SwitchId
        $PortChannelData.Members = $PortChannel | Get-UcsFabricServerPortChannelMember | Select-Object EpDn,PeerDn
        $Data.FabricPcs += $PortChannelData
    }

    # Uplink PortChannels
    $Data.UplinkPcs = @()
    $PortChannels = Get-UcsUplinkPortChannel -Ucs $handle
    foreach ($PortChannel in $PortChannels) {
    # Get-UcsUplinkPortChannel -Ucs $handle | ForEach-Object {
        $PortChannelData = @{}
        $PortChannelData.Name = $PortChannel.Rn
        $PortChannelData.Fabric_Id = $PortChannel.SwitchId
        $PortChannelData.Members = $PortChannel | Get-UcsUplinkPortChannelMember | Select-Object EpDn,PeerDn
        $Data.UplinkPcs += $PortChannelData
    }

    # Qos Domain Policies
    $Data.Qos = @{}
    $Data.Qos.Domain = @()
    $Data.Qos.Domain += Get-UcsQosClass -Ucs $handle | Sort-Object Cos -Descending
    $Data.Qos.Domain += Get-UcsBestEffortQosClass -Ucs $handle
    $Data.Qos.Domain += Get-UcsFcQosClass -Ucs $handle

    # Qos Policies
    $Data.Qos.Policies = @()
    $Policies = Get-UcsQosPolicy -Ucs $handle
    foreach ($Policy in $Policies) {
    # Get-UcsQosPolicy -Ucs $handle | ForEach-Object {
        $PolicyData = @{}
        $PolicyData.Name = $Policy.Name
        $PolicyData.Owner = $Policy.PolicyOwner

        $PolicyDetail = $Policy | Get-UcsChild -ClassId EpqosEgress
        $PolicyData.Burst = $PolicyDetail.Burst
        $PolicyData.HostControl = $PolicyDetail.HostControl
        $PolicyData.Prio = $PolicyDetail.Prio
        $PolicyData.Rate = $PolicyDetail.Rate

        $Data.Qos.Policies += $PolicyData
    }

    # VLANs
    $Data.Vlans = @()
    $Data.Vlans += Get-UcsVlan -Ucs $handle | Where-Object {$_.IfRole -eq "network"} | Sort-Object -Property Ucs,Id

    # Network Control Policies
    $Data.Control_Policies = @()
    $Data.Control_Policies += Get-UcsNetworkControlPolicy -Ucs $handle | Where-Object Dn -ne "fabric/eth-estc/nwctrl-default" | Select-Object Cdp,MacRegisterMode,Name,UplinkFailAction,Descr,Dn,PolicyOwner

    # Mac Address Pool Definitions
    $Data.Mac_Pools = @()
    $Pools = Get-UcsMacPool -Ucs $handle
    foreach ($Pool in $Pools) {
    # Get-UcsMacPool -Ucs $handle | ForEach-Object {
        $PoolData = @{}
        $PoolData.Name = $Pool.Name
        $PoolData.Assigned = $Pool.Assigned
        $PoolData.Size = $Pool.Size
        $MemberBlocks = $Pool | Get-UcsMacMemberBlock -Ucs $handle
        foreach ($MemberBlock in $MemberBlocks) {
            $PoolData.From += $MemberBlock.From
            $PoolData.To += $MemberBlock.To
        }
        # ($PoolData.From,$PoolData.To) = $Pool | Get-UcsMacMemberBlock | Select-Object From,To | ForEach-Object {$_.From,$_.To}
        $Data.Mac_Pools += $PoolData
    }

    # Mac Address Pool Allocations
    $Data.Mac_Allocations = @()
    $Data.Mac_Allocations += Get-UcsMacPoolPooled -Assigned yes | Select-Object Id,Assigned,AssignedToDn

    # Ip Pool Definitions
    $Data.Ip_Pools = @()
    $Pools = Get-UcsIpPool -Ucs $handle
    foreach ($Pool in $Pools) {
    # Get-UcsIpPool -Ucs $handle | ForEach-Object {
        $PoolData = @{}
        $PoolData.Name = $Pool.Name
        $PoolData.Assigned = $Pool.Assigned
        $PoolData.Size = $Pool.Size

        $PoolDetail = $Pool | Get-UcsIpPoolBlock
        $PoolData.From = $PoolDetail.From
        $PoolData.To = $PoolDetail.To
        $PoolData.DefGw = $PoolDetail.DefGw
        $PoolData.Subnet = $PoolDetail.Subnet
        $PoolData.PrimDns = $PoolDetail.PrimDns

        $Data.Ip_Pools += $PoolData
    }

    # Ip Pool Allocations
    $Data.Ip_Allocations = @()
    $Data.Ip_Allocations += Get-UcsIpPoolPooled -Assigned yes | Select-Object AssignedToDn,DefGw,Id,PrimDns,Subnet,Assigned

    # vNic Templates
    $Data.vNic_Templates = @()
    $Data.vNic_Templates += Get-UcsVnicTemplate -Ucs $handle | Select-Object Ucs,Dn,Name,Descr,SwitchId,TemplType,IdentPoolName,Mtu,NwCtrlPolicyName,QosPolicyName

    return $Data
}

function Get-SanTabData {
    <#
    .DESCRIPTION
        Extract SAN tab data: FC/FCOE uplinks and FC/FCOE/NAS storage ports
    .PARAMETER $handle
        Handle (object reference) to target UCS domain
    .PARAMETER $UcsFIs
        List of hashtables containing previously collected FI inventory data (Get-InventoryFIData)
    .PARAMETER $Statistics
        Object reference to Get-UcsStatistics
    .OUTPUTS
    #>

    param (
        [Parameter(Mandatory)]$handle,
        [Parameter(Mandatory)]$UcsFIs,
        [Parameter(Mandatory)]$Statistics
    )

    $Data = @{}
    $Data.UplinkFcoePorts = @()
    $Data.UplinkFcPorts = @()
    $Data.StorageFcPorts = @()
    $Data.StoragePorts = @()

    # Iterate through each FI and grab san performance data based on port role
    foreach ($UcsFI in $UcsFis) {
        # Ethernet NAS, FCOE storage, or FCOE uplink ports
        $Ports = $UcsFI.Ports
        foreach ($Port in $Ports) {
            $PortData = @{}

            $PortData.Dn = $Port.Dn
            $PortData.PortId = $Port.PortId
            $PortData.SlotId = $Port.SlotId
            $PortData.Fabric_Id = $Port.SwitchId
            $PortData.Mac = $Port.Mac
            $PortData.Speed = $Port.OperSpeed
            $PortData.IfType = $Port.IfType
            $PortData.IfRole = $Port.IfRole
            $PortData.XcvrType = $Port.XcvrType
            $PortData.Performance = @{}
            $cmd_args = @{
                UcsStats = $Statistics
                DnFilter = "$($port.Dn)/.*stats"
                StatList = @("TotalBytes","TotalPackets","TotalBytesDeltaAvg")
            }
            $PortData.Performance.Rx = Get-DeviceStats @cmd_args -RnFilter "rx[-]stats"
            $PortData.Performance.Tx = Get-DeviceStats @cmd_args -RnFilter "tx[-]stats"
            $PortData.Status = $Port.OperState
            $PortData.State = $Port.AdminState

            # Store uplinks separately from direct storage ports
            if ($Port.IfRole -cmatch "fc.*uplink") {
                $Data.UplinkFcoePorts += $PortData
            } elseif ($Port.IfRole -cmatch "storage") {
                $Data.StoragePorts += $PortData
            }
        }

        # FC Uplink and Storage Ports
        $Ports = $UcsFI.FcUplinkPorts
        foreach ($Port in $Ports) {
            $PortData = @{}
            $PortData.Dn = $Port.Dn
            $PortData.PortId = $Port.PortId
            $PortData.SlotId = $Port.SlotId
            $PortData.Fabric_Id = $Port.SwitchId
            $PortData.Wwn = $Port.Wwn
            $PortData.IfRole = $Port.IfRole
            $PortData.Speed = $Port.OperSpeed
            $PortData.Mode = $Port.Mode
            $PortData.XcvrType = $Port.XcvrType
            $PortData.Performance = @{}
            $cmd_args = @{
                UcsStats = $Statistics
                DnFilter = "$($port.Dn)/stats"
                RnFilter = "stats"
                StatList = @("BytesRx","PacketsRx","BytesRxDeltaAvg","BytesTx","PacketsTx","BytesTxDeltaAvg")
            }
            $stats = Get-DeviceStats @cmd_args
            $PortData.Performance.Rx = $stats | Select-Object BytesRx,PacketsRx,BytesRxDeltaAvg
            $PortData.Performance.Tx = $stats | Select-Object BytesTx,PacketsTx,BytesTxDeltaAvg
            $PortData.Status = $Port.OperState
            $PortData.State = $Port.AdminState

            # Store uplinks separately from direct storage ports
            if ($Port.IfRole -cmatch "network") {
                $Data.UplinkFcPorts += $PortData
            } elseif ($Port.IfRole -cmatch "storage") {
                $Data.StorageFcPorts += $PortData
            }
        }
    }

    # SAN PortChannel Uplinks
    $Data.UplinkPcs = @()
    $PortChannels = Get-UcsFcUplinkPortChannel -Ucs $handle
    foreach ($PortChannel in $PortChannels) {
        $PortChannelData = @{}
        $PortChannelData.Name = $PortChannel.Rn
        $PortChannelData.Members = $PortChannel | Get-UcsFabricFcSanPcEp | Select-Object EpDn,PeerDn
        $Data.UplinkPcs += $PortChannelData
    }

    # FCoE PortChannel Uplinks
    $Data.FcoePcs = @()
    $PortChannels = Get-UcsFabricFcoeSanPc -Ucs $handle
    foreach ($PortChannel in $PortChannels) {
        $PortChannelData = @{}
        $PortChannelData.Name = $PortChannel.Rn
        $PortChannelData.Members = $PortChannel | Get-UcsFabricFcoeSanPcEp | Select-Object EpDn
        $Data.FcoePcs += $PortChannelData
    }

    # VSANs
    $Data.Vsans = @()
    $Data.Vsans += Get-UcsVsan -Ucs $handle | Select-Object FcoeVlan,Id,name,SwitchId,ZoningState,IfRole,IfType,Transport

    # WWN Pools
    $Data.Wwn_Pools = @()
    $Pools = Get-UcsWwnPool -Ucs $handle
    foreach ($Pool in $Pools) {
        $PoolData = @{}
        $PoolData.Name = $Pool.Name
        $PoolData.Assigned = $Pool.Assigned
        $PoolData.Size = $Pool.Size
        $PoolData.Purpose = $Pool.Purpose
        $PoolData.From = @()
        $PoolData.To = @()
        $MemberBlocks = $Pool | Get-UcsWwnMemberBlock -Ucs $handle
        foreach ($MemberBlock in $MemberBlocks) {
            $PoolData.From += $MemberBlock.From
            $PoolData.To += $MemberBlock.To
        }
        $Data.Wwn_Pools += $PoolData
    }
    # WWN Allocations
    $Data.Wwn_Allocations = @()
    $Data.Wwn_Allocations += Get-UcsWwnInitiator -Assigned yes | Select-Object AssignedToDn,Id,Assigned,Purpose

    # vHba Templates
    $Data.vHba_Templates = Get-UcsVhbaTemplate -Ucs $handle | Select-Object Name,TempType

    return $Data
}

function Get-FaultData {
    <#
    .DESCRIPTION
        Extract faults data
    .PARAMETER $handle
        Handle (object reference) to target UCS domain
    .OUTPUTS
    #>

    param (
        [Parameter(Mandatory)]$handle
    )

    $Data = @()

    $Faults = Get-UcsFault -Ucs $handle -Filter 'Severity -cmatch "critical|major|minor|warning"' | Sort-Object -Property Severity

    # Iterate through each fault and grab information
    foreach ($Fault in $Faults) {
        $FaultData = @{}

        $FaultData.Severity = $Fault.Severity
        $FaultData.Descr = $Fault.Descr
        $FaultData.Dn = $Fault.Dn
        $FaultData.Date = $Fault.Created
        $Data += $FaultData
    }
    return $Data
}

function Invoke-UcsDataGather {
    param (
        [Parameter(Mandatory)]$domain,
        [Parameter(Mandatory)]$Process_Hash
    )
    # Set Job Progress to 0 and connect to the UCS domain passed
    $Process_Hash.Progress[$domain] = 0

    $handle = Connect-Ucs $Process_Hash.Creds[$domain].VIP -Credential $Process_Hash.Creds[$domain].Creds

    # Get all system performance statistics
    if (!$NoStats) {
        $statistics = (Get-UcsStatistics -Ucs $handle).Where({$_.Rn -match "power|temp|vnic|rx|tx-stats|^stats"})
    } else {
        $statistics = ""
    }

    # Get capabilities data from domain embedded catalogs
    $EquipLocalDskDef = Get-UcsEquipmentLocalDiskDef -Ucs $handle
    $EquipManufactDef = Get-UcsEquipmentManufacturingDef -Ucs $handle
    $EquipPhysicalDef = Get-UcsEquipmentPhysicalDef -Ucs $handle

    # Get running firmware for all components
    $AllRunningFirmware = Get-UcsFirmwareRunning -Ucs $handle

    # Get all Maintenance policies
    $MaintenancePolicies = Get-UcsMaintenancePolicy -Ucs $handle

    # Initialize DomainHash variable for this domain
    Start-UcsTransaction -Ucs $handle
    $DomainHash = @{}
    $DomainHash.System = @{}
    $DomainHash.Inventory = @{}
    $DomainHash.Inventory.FIs = @()
    $DomainHash.Inventory.Chassis = @()
    $DomainHash.Inventory.IOMs = @()
    $DomainHash.Inventory.Blades = @()
    $DomainHash.Inventory.Rackmounts = @()
    $DomainHash.Policies = @{}
    $DomainHash.Profiles = @{}
    $DomainHash.Lan = @{}
    $DomainHash.San = @{}
    $DomainHash.Faults = @()

    $DomainHash.ReportDate = Get-Date -format MM/dd/yyyy

    # Get status. Used by system data and FI data
    $DomainStatus = Get-UcsStatus -Ucs $handle
    $DomainName = $DomainStatus.Name
    #===================================#
    #    Start System Data Collection    #
    #===================================#
    # Set Job Progress
    $Process_Hash.Progress[$domain] = 1
    $cmd_args = @{
        handle = $handle
        DomainStatus = $DomainStatus
        Statistics = $Statistics
    }
    $DomainHash.System = Get-SystemData @cmd_args

    #===================================#
    #    Start Inventory Collection        #
    #===================================#

    # Start Fabric Interconnect Inventory Collection

    # Set Job Progress
    $Process_Hash.Progress[$domain] = 12

    $cmd_args = @{
        handle = $handle
        DomainStatus = $DomainStatus
        EquipManufactDef = $EquipManufactDef
    }
    $DomainHash.Inventory.FIs, $SystemItems = Get-InventoryFIData @cmd_args
    # Merge returned System Tab data into existing data
    $DomainHash.System = $DomainHash.System + $SystemItems

    # End Fabric Interconnect Inventory Collection

    # Start Chassis Inventory Collection

    # Set Job Progress
    # $Process_Hash.Progress[$domain] =

    $cmd_args = @{
        handle = $handle
        EquipPhysicalDef = $EquipPhysicalDef
    }
    $DomainHash.Inventory.Chassis += Get-InventoryChassisData @cmd_args

    # End Chassis Inventory Collection

    # Start IOM Inventory Collection

    # Set Job Progress
    $Process_Hash.Progress[$domain] = 24

    $cmd_args = @{
        handle = $handle
        EquipManufactDef = $EquipManufactDef
        AllRunningFirmware = $AllRunningFirmware
    }
    $DomainHash.Inventory.IOMs += Get-InventoryIOModuleData @cmd_args

    # End IOM Inventory Collection

    # Start Server Inventory Collection

    # Get all memory and vif data for future iteration
    $memoryArray = Get-UcsMemoryUnit -Ucs $handle
    $AllFabricEp = Get-UcsFabricPathEp -Ucs $handle

    $cmd_args = @{
        handle = $handle
        AllFabricEp = $AllFabricEp
        memoryArray = $memoryArray
        EquipLocalDskDef = $EquipLocalDskDef
        EquipManufactDef = $EquipManufactDef
        EquipPhysicalDef = $EquipPhysicalDef
        AllRunningFirmware = $AllRunningFirmware
    }

    # Start Blade Inventory Collection
    # Set Job Progress
    $Process_Hash.Progress[$domain] = 36
    $DomainHash.Inventory.Blades += Get-InventoryServerData @cmd_args -IsBlade

    # Start Rack Inventory Collection
    # Set Job Progress
    $Process_Hash.Progress[$domain] = 48

    $DomainHash.Inventory.Rackmounts += Get-InventoryServerData @cmd_args

    # End Server Inventory Collection

    # Start Policy Data and Pools Collection

    # Set Job Progress
    $Process_Hash.Progress[$domain] = 60
    # # Hash variable for storing system policies
    # $DomainHash.Policies.SystemPolicies = @{}

    $cmd_args = @{
        handle = $handle
        MaintenancePolicies = $MaintenancePolicies
    }
    $DomainHash.Policies = Get-PolicyData @cmd_args

    # End Policy Data and Pools Collection

    # Start Service Profile data collection
    # Get Service Profiles by Template

    # Update current job progress
    $Process_Hash.Progress[$domain] = 72
    $cmd_args = @{
        handle = $handle
        MaintenancePolicies = $MaintenancePolicies
        Statistics = $Statistics
    }
    $DomainHash.Profiles = Get-ServiceProfileData @cmd_args

    # End Service Profile Collection

    # Start LAN Configuration
    # Set Job Progress
    # $Process_Hash.Progress[$domain] =

    # Get the collection time interval for port performance
    $DomainHash.Collection = @{}
    $PortCollectionInterval = (Get-UcsCollectionPolicy -Ucs $handle -Name "port").CollectionInterval
    # Normalize collection interval to seconds
    Switch -wildcard (($PortCollectionInterval -split '[0-9]')[-1]) {
        "minute*" {$DomainHash.Collection.Port = ((($PortCollectionInterval -split '[a-z]')[0]) -as [int]) * 60}
        "second*" {$DomainHash.Collection.Port = ((($PortCollectionInterval -split '[a-z]')[0]) -as [int])}
    }

    $cmd_args = @{
        handle = $handle
        UcsFIs = $DomainHash.Inventory.FIs
        Statistics = $Statistics
    }
    $DomainHash.Lan = Get-LanTabData @cmd_args

    # End Lan Configuration

    # Start SAN Configuration
    # Set Job Progress
    # $Process_Hash.Progress[$domain] =

    $cmd_args = @{
        handle = $handle
        UcsFIs = $DomainHash.Inventory.FIs
        Statistics = $Statistics
    }
    $DomainHash.San = Get-SanTabData @cmd_args

    # End San Configuration

    # Start Fault List

    # Set Job Progress
    $Process_Hash.Progress[$domain] = 84

    $DomainHash.Faults += Get-FaultData -handle $handle

    # End Fault List

    # Set Job Progress
    $Process_Hash.Progress[$domain] = 96

    Complete-UcsTransaction -Ucs $handle
    # Add current Domain data to global process Hash
    $Process_Hash.Domains[$DomainName] = $DomainHash
    # Disconnect from current UCS domain
    Disconnect-Ucs -Ucs $handle
}

function Start-UcsDataGather {
    <#
    .DESCRIPTION
        Function for creating the html configuration report for all of the connected UCS domains
    #>
    # Check to ensure an active UCS handle exists before generating the report
    if(!(Confirm-AnyUcsHandle)) {
        Read-Host "There are currently no connected UCS domains`n`nPress any key to continue"
        return
    }

    # Grab filename for the report
    $OutputFile = Get-SaveFile($dflt_output_path)

    Clear-Host
    Write-Host "Generating Report..."

    # Get Start time to track report generation run time
    $start_timestamp = Get-Date

    # Creates a synchronized hash variable of all the UCS domains and credential info
    $Process_Hash = [hashtable]::Synchronized(@{})
    $Process_Hash.Creds = $UCS_Creds
    $Process_Hash.Keys = @()
    $Process_Hash.Keys += $UCS.get_keys()
    $Process_Hash.Domains = @{}
    $Process_Hash.Progress = @{}

    $recurse_script = $MyInvocation.ScriptName

    # Initialize runspaces to allow simultaneous domain data collection (could also use workflows)
    $Script:runspaces = New-Object System.Collections.ArrayList
    $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
    $runspacepool = [runspacefactory]::CreateRunspacePool(1, 10, $sessionstate, $Host)
    $runspacepool.Open()
    # Iterate through each domain key and pass to GetUcsData script block
    $UCS.get_keys() | ForEach-Object {
        $domain = $_
        # Create powershell thread to execute script block with current domain
        $powershell = [powershell]::Create().
                        AddCommand($recurse_script).
                            AddParameter("RunAsProcess").
                            AddParameter("NoStats",$NoStats).
                            AddParameter("Domain",$domain).
                            AddParameter("ProcessHash",$Process_Hash)
        $powershell.RunspacePool = $runspacepool
        $temp = "" | Select-Object PowerShell,Runspace,Computer
        $Temp.Computer = $Computer
        $temp.PowerShell = $powershell
        # Invoke runspace
        $temp.Runspace = $powershell.BeginInvoke()
        Write-Verbose ("Adding {0} collection" -f $temp.Computer)
        $runspaces.Add($temp) | Out-Null
    }

    Do {
        # Monitor each job progress and update write-progress
        $Progress = 0
        # Catch conditions where no script progress has occurred
        try {
            # Iterate through each process and add progress divided by process count
            $Process_Hash.Progress.GetEnumerator() | ForEach-Object {
                $Progress += ($_.Value / $Process_Hash.Progress.Count)
            }
        }
        catch {}

        # Write Progress to alert user of overall progress
        $cmd_args = @{
            Activity = "Collecting Report Data. This will take several minutes."
            Status = "$($Progress)% Complete"
            PercentComplete = $Progress
            CurrentOperation = "Additional info TBD."
        }
        Write-Progress @cmd_args

        $more = $false

        # Iterate through each runspace in progress
        Foreach($runspace in $runspaces) {
            # If runspace is complete cleanly end/exit runspace
            If ($runspace.Runspace.isCompleted) {
                $runspace.powershell.EndInvoke($runspace.Runspace)
                $runspace.powershell.dispose()
                $runspace.Runspace = $null
                $runspace.powershell = $null
            } ElseIf ($null -ne $runspace.Runspace) {
                $more = $true
            }
        }

        # Sleep for 100ms before updating progress
        If ($more) {Start-Sleep -Milliseconds 100}

        # Clean out unused runspace jobs
        $temphash = $runspaces.clone()
        $temphash | Where-Object {$Null -eq $_.runspace} | ForEach-Object {
            Write-Verbose ("Removing {0}" -f $_.computer)
            $Runspaces.remove($_)
        }
        [console]::Title = ("Remaining Runspace Jobs: {0}" -f ((@($runspaces | Where-Object {$Null -ne $_.Runspace}).Count)))
    } while ($more)

    # Update overall progress to complete
    Write-Progress "Done" "Done" -Completed

    # End collection script

    # Start HTML report generation
    # Import template file to memory
    $ReportRawText = Get-Content ./Report_Template.htm

    # Convert hash table to JSON
    $ReportData = $Process_Hash.Domains | ConvertTo-JSON -Depth 14 -Compress

    # Replace date and JSON placeholders with real data
    $ReportRawText = $ReportRawText.Replace('DATE_REPLACE',(Get-Date -format MM/dd/yyyy))
    $ReportRawText = $ReportRawText.Replace('{"_comment": "placeholder"}', $ReportData)

    # Save configuration report to file
    Set-Content -Path $OutputFile -Value $ReportRawText

    # Email Report if Email switch is set or Email_Report is set
    if ($Email) {$mail_to = $Email} else {$mail_to = $test_mail_to}

    if ($test_mail_flag -or $Email) {
        if ($mail_server -and $mail_from -and $mail_to) {
            $cmd_args = @{
                To = $mail_to
                From = $mail_from
                SmtpServer = $mail_server
                Subject = "Cisco UCS Configuration Report"
                Body = "Hi-`n`nThis is an automatic message by UCS Configuration Report.  Please see attachment."
                Attachment = $OutputFile
            }
            Send-MailMessage @cmd_args -UseSsl
        } else {
            Write-Host "Not all email parameters configured."
        }
    }

    # Write elapsed time to user
    Write-host "Total Elapsed Time: $(Get-ElapsedTime -FirstTimestamp $start_timestamp)`n"
    if(-Not $Silent) {Read-Host "UCS Configuration Report has been generated.  Press any key to continue..."}
}

function Disconnect-AllUcsDomains {
    <#
    .DESCRIPTION
        Disconnects all UCS Domains and exits the script
    #>
    foreach ($Domain in $UCS.get_keys()) {
        if(Test-UcsHandle($UCS[$Domain])) {
            Write-Host "Disconnecting $($UCS[$Domain].Name)..."
            Disconnect-Ucs -Ucs $UCS[$Domain].Handle
            $script:UCS[$Domain].Remove("Handle")
        }
    }
    $script:UCS = $null
    Write-Host "Exiting Program`n"
}

function Test-RequiredPsModules {
    <#
    .DESCRIPTION
        Function that checks that all required powershell modules are present
    #>
    if(@(Get-Module -ListAvailable -Name Cisco.UCSManager).Count -lt 1) {
        # Module not installed.
        Write-Host "UCS POWERTOOLS NOT DETECTED! Please perform installation instructions listed in README.md."
        exit
    }
}

# === Main ===
Start-Main


