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
        [Switch]$IsBlade
    )
    $AllRunningFirmware = Get-UcsFirmwareRunning
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
            $ServerData.Rack_Id = $rack.Id
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
            $AdapterData.Fw = ($Adapter | Get-UcsMgmtController | Get-UcsFirmwareRunning -Deployment system).Version
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
            $Disks = $Controller | Get-UcsStorageLocalDisk -Presence "equipped"
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
        Extract inventory data for blades or rackmount servers
    .PARAMETER $handle
        Handle (object reference) to target UCS domain
    .PARAMETER $DomainStatus
        Object reference to results of Get-UcsStatus
    .OUTPUTS
        Hashtable
    #>

    param (
        [Parameter(Mandatory)]$handle,
        [Parameter(Mandatory)]$DomainStatus
    )
    $Data = @{}
    $Data.Chassis_Power = @()
    $Data.Server_Power = @()
    $Data.Server_Temp = @()

    # Get UCS Cluster State
    $Data.VIP = $DomainStatus.VirtualIpv4Address
    $Data.UCSM = (Get-UcsMgmtController -Ucs $handle -Subject system | Get-UcsFirmwareRunning).Version
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

    # Initialize DomainHash variable for this domain
    Start-UcsTransaction -Ucs $handle
    $DomainHash = @{}
    $DomainHash.System = @{}
    $DomainHash.Inventory = @{}
    $DomainHash.Inventory.Blades = @()
    $DomainHash.Inventory.Rackmounts = @()
    $DomainHash.Policies = @{}
    $DomainHash.Profiles = @{}
    $DomainHash.Lan = @{}
    $DomainHash.San = @{}
    $DomainHash.Faults = @()

    # Get status. Used by system data and FI data
    $DomainStatus = Get-UcsStatus -Ucs $handle
    $DomainName = $DomainStatus.Name
    #===================================#
    #    Start System Data Collection    #
    #===================================#
    $Process_Hash.Progress[$domain] = 1
    $cmd_args = @{
        handle = $handle
        DomainStatus = $DomainStatus
    }
    $DomainHash.System = Get-SystemData @cmd_args

    #===================================#
    #    Start Inventory Collection        #
    #===================================#

    # Start Fabric Interconnect Inventory Collection

    # Set Job Progress
    $Process_Hash.Progress[$domain] = 12
    $DomainHash.Inventory.FIs = @()
    # Iterate through Fabric Interconnects and grab relevant data points
    Get-UcsNetworkElement -Ucs $handle | ForEach-Object {
        # Store current pipe value to fi variable
        $fi = $_
        # Hash variable for storing current FI details
        $fiHash = @{}
        $fiHash.Dn = $fi.Dn
        $fiHash.Fabric_Id = $fi.Id
        $fiHash.Operability = $fi.Operability
        $fiHash.Thermal = $fi.Thermal
        # Get leadership role and management service state
        if($fi.Id -eq "A") {
            $fiHash.Role = $DomainStatus.FiALeadership
            $fiHash.State = $DomainStatus.FiAManagementServicesState
        } else {
            $fiHash.Role = $DomainStatus.FiBLeadership
            $fiHash.State = $DomainStatus.FiAManagementServicesState
        }

        # Get the common name of the fi from the manufacturing definition and format the text
        $fiModel = ($EquipManufactDef | Where-Object  {$_.Sku -cmatch $($fi.Model)} | Select-Object Name).Name -replace "Cisco UCS ", ""
        if($fiModel -is [array]) {$fiHash.Model = $fiModel.Item(0) -replace "Cisco UCS ", ""} else {$fiHash.Model = $fiModel -replace "Cisco UCS ", ""}

        $fiHash.Serial = $fi.Serial
        # Get FI System and Kernel FW versions
        ${fiBoot} = Get-UcsMgmtController -Ucs $handle -Dn "$($fi.Dn)/mgmt" | Get-ucsfirmwarebootdefinition | Get-UcsFirmwareBootUnit -Filter 'Type -ieq system -or Type -ieq kernel' | Select-Object Type,Version
        $fiHash.System = (${fiBoot} | Where-Object {$_.Type -eq "system"}).Version
        $fiHash.Kernel = (${fiBoot} | Where-Object {$_.Type -eq "kernel"}).Version

        # Get out of band management IP and Port licensing information
        $fiHash.IP = $fi.OobIfIp
        $ucsLicense = Get-UcsLicense -Ucs $handle -Scope $fi.Id
        $ports_used = ($ucsLicense | Select-Object UsedQuant).UsedQuant
        if ($ports_used -is [system.array]) {
            $ports_used_total = 0
            $ports_used | ForEach-Object {$ports_used_total += $_}
            $fiHash.Ports_Used = $ports_used_total
            Remove-Variable ports_used_total
        } else {
            $fiHash.Ports_Used = $ports_used
        }
        $ports_used_sub = ($ucsLicense | Select-Object SubordinateUsedQuant).SubordinateUsedQuant
        if ($ports_used_sub -and $ports_used_sub -is [system.array]) {
            $ports_used_sub_total = 0
            $ports_used_sub | ForEach-Object {$ports_used_sub_total += $_}
            $fiHash.Ports_Used += $ports_used_sub_total
            Remove-Variable ports_used_sub_total
        } else {
            $fiHash.Ports_Used += $ports_used_sub
        }
        Remove-Variable ports_used_sub
        Remove-Variable ports_used
        $ports_licensed = ($ucsLicense | Select-Object AbsQuant).AbsQuant
        if ($ports_licensed -is [system.array]) {
            $ports_licensed_total = 0
            $ports_licensed | ForEach-Object {$ports_licensed_total += $_}
            $fiHash.Ports_Licensed = $ports_licensed_total
            Remove-Variable ports_licensed_total
        } else {
            $fiHash.Ports_Licensed = $ports_licensed
        }
        Remove-Variable ports_licensed
        if ($ports_used) {Remove-Variable ports_used}

        # Get Ethernet and FC Switching mode of FI
        $fiHash.Ethernet_Mode = (Get-UcsLanCloud -Ucs $handle).Mode
        $fiHash.FC_Mode = (Get-UcsSanCloud -Ucs $handle).Mode

        # Get Local storage, VLAN, and Zone utilization numbers
        $fiHash.Storage = $fi | Get-UcsStorageItem | Select-Object Name,Size,Used
        $fiHash.VLAN = $fi | Get-UcsSwVlanPortNs | Select-Object Limit,AccessVlanPortCount,BorderVlanPortCount,AllocStatus
        $fiHash.Zone = $fi | Get-UcsManagedObject -Classid SwFabricZoneNs | Select-Object Limit,ZoneCount,AllocStatus

        # Sort Expression to filter port id to be just the numerical port number and sort ascending
        $sortExpr = {if ($_.Dn -match "(?=port[-]).*") {($matches[0] -replace ".*(?<=[-])",'') -as [int]}}
        # Get Fabric Port Configuration and sort by port id using the above sort expression
        $fiHash.Ports = Get-UcsFabricPort -Ucs $handle -SwitchId "$($fi.Id)" -AdminState enabled | Sort-Object $sortExpr | Select-Object AdminState,Dn,IfRole,IfType,LicState,LicGP,Mac,Mode,OperState,OperSpeed,XcvrType,PeerDn,PeerPortId,PeerSlotId,PortId,SlotId,SwitchId
        $fiHash.FcUplinkPorts = Get-UcsFiFcPort -Ucs $handle -SwitchId "$($fi.Id)" -AdminState 'enabled' | Sort-Object $sortExpr | Select-Object AdminState,Dn,IfRole,IfType,LicState,LicGP,Wwn,Mode,OperState,OperSpeed,XcvrType,PortId,SlotId,SwitchId

        # Store fi hash to domain hash variable
        $DomainHash.Inventory.FIs += $fiHash

        # Get FI Role and IP for system tab of report
        if($fiHash.Fabric_Id -eq 'A') {
            $DomainHash.System.FI_A_Role = $fiHash.Role
            $DomainHash.System.FI_A_IP = $fiHash.IP
        } else {
            $DomainHash.System.FI_B_Role = $fiHash.Role
            $DomainHash.System.FI_B_IP = $fiHash.IP
        }

    }
    # End FI Inventory Collection

    # Start Chassis Inventory Collection

    # Initialize array variable for storing Chassis data
    $DomainHash.Inventory.Chassis = @()
    # Iterate through chassis inventory and grab relevant data
    Get-UcsChassis -Ucs $handle | ForEach-Object {
        # Store current pipe variable
        $chassis = $_
        # Hash variable for storing current chassis data
        $chassisHash = @{}
        $chassisHash.Dn = $chassis.Dn
        $chassisHash.Id = $chassis.Id
        $chassisHash.Model = $chassis.Model
        $chassisHash.Status = $chassis.OperState
        $chassisHash.Operability = $chassis.Operability
        $chassisHash.Power = $chassis.Power
        $chassisHash.Thermal = $chassis.Thermal
        $chassisHash.Serial = $chassis.Serial
        $chassisHash.Blades = @()

        # Initialize chassis used slot count to 0
        $slotCount = 0
        # Iterate through all blades within current chassis
        $chassis | Get-UcsBlade | Select-Object Model,SlotId,AssignedToDn | ForEach-Object {
            # Hash variable for storing current blade data
            $bladeHash = @{}
            $bladeHash.Model = $_.Model
            $bladeHash.SlotId = $_.SlotId
            $bladeHash.Service_Profile = $_.AssignedToDn
            # Get width of blade and convert to slot count
            $bladeHash.Width = [math]::floor((($EquipPhysicalDef | Where-Object {$_.Dn -ilike "*$($bladeHash.Model)*"}).Width)/8)
            # Increment used slot count by current blade width
            $slotCount += $bladeHash.Width
            $chassisHash.Blades += $bladeHash
        }
        # Get Used slots and slots available from iterated slot count
        $chassisHash.SlotsUsed = $slotCount
        $chassisHash.SlotsAvailable = 8 - $slotCount

        # Get chassis PSU data and redundancy mode
        $chassisHash.Psus = @()
        $chassisHash.Psus = $chassis | Get-UcsPsu | Sort-Object Id | Select-Object Type,Id,Model,Serial,Dn
        $chassisHash.Power_Redundancy = ($chassis | Get-UcsComputePsuControl | Select-Object Redundancy).Redundancy

        # Add chassis to domain hash variable
        $DomainHash.Inventory.Chassis += $chassisHash
    }
    # End Chassis Inventory Collection

    # Start IOM Inventory Collection

    # Increment job progress
    $Process_Hash.Progress[$domain] = 24

    # Get all fabric and blackplane ports for future iteration
    $FabricPorts = Get-UcsEtherSwitchIntFIo -Ucs $handle
    $BackplanePorts = Get-UcsEtherServerIntFIo -Ucs $handle

    # Initialize array for storing IOM inventory data
    $DomainHash.Inventory.IOMs = @()
    # Iterate through each IOM and grab relevant data
    Get-UcsIom -Ucs $handle | Select-Object ChassisId,SwitchId,Model,Serial,Dn | ForEach-Object {
        $iom = $_
        $iomHash = @{}
        $iomHash.Dn = $iom.Dn
        $iomHash.Chassis = $iom.ChassisId
        $iomHash.Fabric_Id = $iom.SwitchId

        # Get common name of IOM model and format for viewing
        $iomHash.Model = ($EquipManufactDef | Where-Object {$_.Sku -cmatch $($iom.Model)}).Name -replace "Cisco UCS ", ""
        $iomHash.Serial = $iom.Serial

        # Get the IOM uplink port channel name if configured
        $iomHash.Channel = (Get-ucsportgroup -Ucs $handle -Dn "$($iom.Dn)/fabric-pc" | Get-UcsEtherSwitchIntFIoPc).Rn

        # Get IOM running and backup fw versions
        $iomHash.Running_FW = (Get-UcsMgmtController -Ucs $handle -Dn "$($iom.Dn)/mgmt" | Get-UcsFirmwareRunning -Deployment system | Select-Object Version).Version
        $iomHash.Backup_FW = (Get-UcsMgmtController -Ucs $handle -Dn "$($iom.Dn)/mgmt" | Get-UcsFirmwareUpdatable | Select-Object Version).Version

        # Initialize FabricPorts array for storing IOM port data
        $iomHash.FabricPorts = @()

        # Iterate through all fabric ports tied to the current IOM
        $FabricPorts | Where-Object {$_.ChassisId -eq "$($iomHash.Chassis)" -and $_.SwitchId -eq "$($iomHash.Fabric_Id)"} | Select-Object SlotId,PortId,OperState,EpDn,PeerSlotId,PeerPortId,SwitchId,Ack,PeerDn | ForEach-Object {
            # Hash variable for storing current fabric port data
            $portHash = @{}
            $portHash.Name = 'Fabric Port ' + $_.SlotId + '/' + $_.PortId
            $portHash.OperState = $_.OperState
            $portHash.PortChannel = $_.EpDn
            $portHash.PeerSlotId = $_.PeerSlotId
            $portHash.PeerPortId = $_.PeerPortId
            $portHash.FabricId = $_.SwitchId
            $portHash.Ack = $_.Ack
            $portHash.Peer = $_.PeerDn
            # Add current fabric port hash variable to FabricPorts array
            $iomHash.FabricPorts += $portHash
        }
        # Initialize BackplanePorts array for storing IOM port data
        $iomHash.BackplanePorts = @()

        # Iterate through all backplane ports tied to the current IOM
        $BackplanePorts | Where-Object {$_.ChassisId -eq "$($iomHash.Chassis)" -and $_.SwitchId -eq "$($iomHash.Fabric_Id)"} | Sort-Object {($_.SlotId) -as [int]},{($_.PortId) -as [int]} | Select-Object SlotId,PortId,OperState,EpDn,SwitchId,PeerDn | ForEach-Object {
            # Hash variable for storing current backplane port data
            $portHash = @{}
            $portHash.Name = 'Backplane Port ' + $_.SlotId + '/' + $_.PortId
            $portHash.OperState = $_.OperState
            $portHash.PortChannel = $_.EpDn
            $portHash.FabricId = $_.SwitchId
            $portHash.Peer = $_.PeerDn
            # Add current backplane port hash variable to FabricPorts array
            $iomHash.BackplanePorts += $portHash
        }
        # Add IOM to domain hash variable
        $DomainHash.Inventory.IOMs += $iomHash
    }
    # End IOM Inventory Collection

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
    }
    # Start Blade Inventory Collection

    # Set progress of current job
    $Process_Hash.Progress[$domain] = 36

    $DomainHash.Inventory.Blades += Get-InventoryServerData @cmd_args -IsBlade

    # Start Rack Inventory Collection
    # Set progress of current job
    $Process_Hash.Progress[$domain] = 48
    $DomainHash.Inventory.Rackmounts += Get-InventoryServerData @cmd_args


    # Start Policy Data Collection

    # Update job progress percent
    $Process_Hash.Progress[$domain] = 60
    # Hash variable for storing system policies
    $DomainHash.Policies.SystemPolicies = @{}
    # Grab DNS and NTP data
    $DomainHash.Policies.SystemPolicies.DNS = @()
    $DomainHash.Policies.SystemPolicies.DNS += (Get-UcsDnsServer -Ucs $handle).Name | Select-Object -Unique
    $DomainHash.Policies.SystemPolicies.NTP = @()
    $DomainHash.Policies.SystemPolicies.NTP += (Get-UcsNtpServer -Ucs $handle).Name | Select-Object -Unique
    # Get chassis discovery data for future use
    $Chassis_Discovery = Get-UcsChassisDiscoveryPolicy -Ucs $handle | Select-Object Action,LinkAggregationPref
    $DomainHash.Policies.SystemPolicies.Action = $Chassis_Discovery.Action
    $DomainHash.Policies.SystemPolicies.Grouping = $Chassis_Discovery.LinkAggregationPref
    $DomainHash.Policies.SystemPolicies.Power = (Get-UcsPowerControlPolicy -Ucs $handle | Select-Object Redundancy).Redundancy
    $DomainHash.Policies.SystemPolicies.FirmwareAutoSyncAction = (Get-UcsFirmwareAutoSyncPolicy | Select-Object SyncState).SyncState
    $DomainHash.Policies.SystemPolicies.Maint = (Get-UcsMaintenancePolicy -Name "default" -Ucs $handle | Select-Object UptimeDisr).UptimeDisr
    $DomainHash.Policies.SystemPolicies.Timezone = (Get-UcsTimezone -Ucs $handle).Timezone | Select-Object -Unique

    # Maintenance Policies
    $DomainHash.Policies.Maintenance = @()
    $DomainHash.Policies.Maintenance += Get-UcsMaintenancePolicy -Ucs $handle | Select-Object Name,Dn,UptimeDisr,Descr,SchedName

    # Host Firmware Packages
    $DomainHash.Policies.FW_Packages = @()
    $DomainHash.Policies.FW_Packages += Get-UcsFirmwareComputeHostPack -Ucs $handle | Select-Object Name,BladeBundleVersion,RackBundleVersion

    # LDAP Policy Data
    $DomainHash.Policies.LDAP_Providers = @()
    $DomainHash.Policies.LDAP_Providers += Get-UcsLdapProvider -Ucs $handle | Select-Object Name,Rootdn,Basedn,Attribute
    $mappingArray = @()
    $DomainHash.Policies.LDAP_Mappings = @()
    $mappingArray += Get-UcsLdapGroupMap -Ucs $handle
    $mappingArray | ForEach-Object {
        $mapHash = @{}
        $mapHash.Name = $_.Name
        $mapHash.Roles = ($_ | Get-UcsUserRole).Name
        $mapHash.Locales = ($_ | Get-UcsUserLocale).Name
        $DomainHash.Policies.LDAP_Mappings += $mapHash
    }

    # Boot Order Policies
    $DomainHash.Policies.Boot_Policies = @()
    $AllBootPolicies = Get-UcsBootPolicy -Ucs $handle
    $DomainHash.Policies.Boot_Policies += Get-ConfiguredBootOrder -BootPolicies $AllBootPolicies

    # End System Policies Collection

    # Start ID Pool Collection
    # External Mgmt IP Pool
    $DomainHash.Policies.Mgmt_IP_Pool = @{}
    # Get the default external management pool
    $mgmtPool = Get-ucsippoolblock -Ucs $handle -Filter "Dn -cmatch ext-mgmt"
    $DomainHash.Policies.Mgmt_IP_Pool.From = $mgmtPool.From
    $DomainHash.Policies.Mgmt_IP_Pool.To = $mgmtPool.To
    $parentPool = $mgmtPool | get-UcsParent
    $DomainHash.Policies.Mgmt_IP_Pool.Size = $parentPool.Size
    $DomainHash.Policies.Mgmt_IP_Pool.Assigned = $parentPool.Assigned

    # Mgmt IP Allocation
    $DomainHash.Policies.Mgmt_IP_Allocation = @()
    $parentPool | Get-UcsIpPoolPooled -Filter "Assigned -ieq yes" | Select-Object AssignedToDn,Id,Subnet,DefGw | ForEach-Object {
        $allocationHash = @{}
        $allocationHash.Dn = $_.AssignedToDn -replace "/mgmt/*.*", ""
        $allocationHash.IP = $_.Id
        $allocationHash.Subnet = $_.Subnet
        $allocationHash.GW = $_.DefGw
        $DomainHash.Policies.Mgmt_IP_Allocation += $allocationHash
    }
    # UUID
    $DomainHash.Policies.UUID_Pools = @()
    $DomainHash.Policies.UUID_Pools += Get-UcsUuidSuffixPool -Ucs $handle | Select-Object Dn,Name,AssignmentOrder,Prefix,Size,Assigned
    $DomainHash.Policies.UUID_Assignments = @()
    $DomainHash.Policies.UUID_Assignments += Get-UcsUuidpoolAddr -Ucs $handle -Assigned yes | select-object AssignedToDn,Id | sort-object -property AssignedToDn

    # Server Pools
    $DomainHash.Policies.Server_Pools = @()
    $DomainHash.Policies.Server_Pools += Get-UcsServerPool -Ucs $handle | Select-Object Dn,Name,Size,Assigned
    $DomainHash.Policies.Server_Pool_Assignments = @()
    $DomainHash.Policies.Server_Pool_Assignments += Get-UcsServerPoolAssignment -Ucs $handle | Select-Object Name,AssignedToDn

    # End ID Pools Collection

    # Start Service Profile data collection
    # Get Service Profiles by Template

    # Update current job progress
    $Process_Hash.Progress[$domain] = 72
    # Grab all Service Profiles
    $profiles = Get-ucsServiceProfile -Ucs $handle

    # Array variable for storing template data
    $templates = @()
    # Grab all service profile templates
    $templates += ($profiles | Where-Object {$_.Type -match "updating[-]template|initial[-]template"} | Select-Object Dn).Dn
    # Add an empty template entry for profiles not bound to a template
    $templates += ""
    # Iterate through templates and grab configuration data
    $templates | ForEach-Object {
        # Grab the current template name
        $templateDn = $_
        $templateId = $templateDn #-replace "/",":"
        # Unchanged copy of the current template name used later in the script
        # Find the profile template that matches the current name
        $template = $profiles | Where-Object {$_.Dn -eq "$templateDn"}
        $templateName = If ($template) {$template.Name} Else {"Unbound"}
        # Hash variable to store data for current templateName
        $DomainHash.Profiles[$templateId] = @{}
        # Switch statement to format the template type
        switch ($template.Type) {
                "updating-template"    {$DomainHash.Profiles[$templateId].Type = "Updating"}
                "initial-template"    {$DomainHash.Profiles[$templateId].Type = "Initial"}
                default {$DomainHash.Profiles[$templateId].Type = "N/A"}
        }
        # Template Details - General Tab

        # Hash variable for storing general template data
        $DomainHash.Profiles[$templateId].General = @{}
        $DomainHash.Profiles[$templateId].General.Name = $templateName
        $DomainHash.Profiles[$templateId].General.Type = $DomainHash.Profiles[$templateId].Type
        $DomainHash.Profiles[$templateId].General.Description = $template.Descr
        $DomainHash.Profiles[$templateId].General.UUIDPool = $template.IdentPoolName
        $DomainHash.Profiles[$templateId].General.Boot_Policy = $template.OperBootPolicyName
        $DomainHash.Profiles[$templateId].General.PowerState = ($template | Get-UcsServerPower).State
        $DomainHash.Profiles[$templateId].General.MgmtAccessPolicy = $template.ExtIPState
        $DomainHash.Profiles[$templateId].General.Server_Pool = $template | Get-UcsServerPoolAssignment | Select-Object Name,Qualifier,RestrictMigration
        if ($templateDn -eq "") {$DomainHash.Profiles[$templateId].General.Maintenance_Policy = ""} else {$DomainHash.Profiles[$templateId].General.Maintenance_Policy = Get-UcsMaintenancePolicy -Ucs $handle -Filter "Dn -ieq $($template.OperMaintPolicyName)" | Select-Object Name,Dn,Descr,UptimeDisr}

        # Template Details - Storage Tab

        # Hash variable for storing storage template data
        $DomainHash.Profiles[$templateId].Storage = @{}

        # Node WWN Configuration
        $fcNode = $template | Get-UcsVnicFcNode
        # Grab VNIC connectivity
        $vnicConn = $template | Get-UcsVnicConnDef
        $DomainHash.Profiles[$templateId].Storage.Nwwn = $fcNode.Addr
        $DomainHash.Profiles[$templateId].Storage.Nwwn_Pool = $fcNode.IdentPoolName
        if ($templateDn -eq "") {$DomainHash.Profiles[$templateId].Storage.Local_Disk_Config = ""} else {$DomainHash.Profiles[$templateId].Storage.Local_Disk_Config = Get-UcsLocalDiskConfigPolicy -Dn $template.OperLocalDiskPolicyName | Select-Object Mode,ProtectConfig,XtraProperty}
        $DomainHash.Profiles[$templateId].Storage.Connectivity_Policy = $vnicConn.SanConnPolicyName
        $DomainHash.Profiles[$templateId].Storage.Connectivity_Instance = $vnicConn.OperSanConnPolicyName
        # Array variable for storing HBA data
        $DomainHash.Profiles[$templateId].Storage.Hbas = @()
        $template | Get-UcsVhba | ForEach-Object {
            $hbaHash = @{}
            $hbaHash.Name = $_.Name
            $hbaHash.Pwwn = $_.IdentPoolName
            $hbaHash.FabricId = $_.SwitchId
            $hbaHash.Desired_Order = $_.Order
            $hbaHash.Actual_Order = $_.OperOrder
            $hbaHash.Desired_Placement = $_.AdminVcon
            $hbaHash.Actual_Placement = $_.OperVcon
            $hbaHash.Vsan = ($_ | Get-UcsChild | Select-Object Name).Name
            $DomainHash.Profiles[$templateId].Storage.Hbas += $hbaHash
        }

        # Template Details - Network Tab

        # Hash variable for storing template network configuration
        $DomainHash.Profiles[$templateId].Network = @{}
        # Lan Connectivity Policy
        $DomainHash.Profiles[$templateId].Network.Connectivity_Policy = $vnicConn.LanConnPolicyName
        $DomainHash.Profiles[$templateId].Network.DynamicVnic_Policy = $template.DynamicConPolicyName
        # Array variable for storing
        $DomainHash.Profiles[$templateId].Network.Nics = @()
        # Iterate through each NIC and grab configuration details
        $template | Get-UcsVnic | ForEach-Object {
            $nicHash = @{}
            $nicHash.Name = $_.Name
            $nicHash.Mac_Address = $_.Addr
            $nicHash.Desired_Order = $_.Order
            $nicHash.Actual_Order = $_.OperOrder
            $nicHash.Fabric_Id = $_.SwitchId
            $nicHash.Desired_Placement = $_.AdminVcon
            $nicHash.Actual_Placement = $_.OperVcon
            $nicHash.Adaptor_Profile = $_.AdaptorProfileName
            $nicHash.Control_Policy = $_.NwCtrlPolicyName
            # Array for storing VLANs
            $nicHash.Vlans = @()
            # Grab all VLANs
            $nicHash.Vlans += $_ | Get-UcsChild -ClassId VnicEtherIf | Select-Object OperVnetName,Vnet,DefaultNet | Sort-Object {($_.Vnet) -as [int]}
            $DomainHash.Profiles[$templateId].Network.Nics += $nicHash
        }

        # Template Details - iSCSI vNICs Tab

        # Array variable for storing iSCSI configuration
        $DomainHash.Profiles[$templateId].iSCSI = @()
        # Iterate through iSCSI interface configuration
        $template | Get-UcsVnicIscsi | ForEach-Object {
            $iscsiHash = @{}
            $iscsiHash.Name = $_.Name
            $iscsiHash.Overlay = $_.VnicName
            $iscsiHash.Iqn = $_.InitiatorName
            $iscsiHash.Adapter_Policy = $_.AdaptorProfileName
            $iscsiHash.Mac = $_.Addr
            $iscsiHash.Vlan = ($_ | Get-UcsVnicVlan).VlanName
            $DomainHash.Profiles[$templateId].iSCSI += $iscsiHash
        }

        # Template Details - Policies Tab

        # Hash variable for storing template Policy configuration data
        $DomainHash.Profiles[$templateId].Policies = @{}
        $DomainHash.Profiles[$templateId].Policies.Bios = $template.BiosProfileName
        $DomainHash.Profiles[$templateId].Policies.Fw = $template.HostFwPolicyName
        $DomainHash.Profiles[$templateId].Policies.Ipmi = $template.MgmtAccessPolicyName
        $DomainHash.Profiles[$templateId].Policies.Power = $template.PowerPolicyName
        $DomainHash.Profiles[$templateId].Policies.Scrub = $template.ScrubPolicyName
        $DomainHash.Profiles[$templateId].Policies.Sol = $template.SolPolicyName
        $DomainHash.Profiles[$templateId].Policies.Stats = $template.StatsPolicyName

        # Service Profile Instances

        # Array variable for storing profiles tied to current template name
        $DomainHash.Profiles[$templateId].Profiles = @()
        # Iterate through all profiles tied to the current template name
        $profiles | Where-Object {$_.OperSrcTemplName -ieq "$templateDn" -and $_.Type -ieq "instance"} | ForEach-Object {
            # Store current pipe variable to local variable
            $sp = $_
            # Hash variable for storing current profile configuration data
            $profileHash = @{}
            $profileHash.Dn = $sp.Dn
            $profileHash.Service_Profile = $sp.Name
            $profileHash.UsrLbl = $sp.UsrLbl
            $profileHash.Assigned_Server = $sp.PnDn
            $profileHash.Assoc_State = $sp.AssocState
            $profileHash.Maint_Policy = $sp.MaintPolicyName
            $profileHash.Maint_PolicyInstance = $sp.OperMaintPolicyName
            $profileHash.FW_Policy = $sp.HostFwPolicyName
            $profileHash.BIOS_Policy = $sp.BiosProfileName
            $profileHash.Boot_Policy = $sp.OperBootPolicyName

            # Service Profile Details - General Tab

            # Hash variable for storing general profile configuration data
            $profileHash.General = @{}
            $profileHash.General.Name = $sp.Name
            $profileHash.General.Overall_Status = $sp.operState
            $profileHash.General.AssignState = $sp.AssignState
            $profileHash.General.AssocState = $sp.AssocState
            $profileHash.General.Power_State = ($sp | Get-UcsChild -ClassId LsPower | Select-Object State).State

            $profileHash.General.UserLabel = $sp.UsrLbl
            $profileHash.General.Descr = $sp.Descr
            $profileHash.General.Owner = $sp.PolicyOwner
            $profileHash.General.Uuid = $sp.Uuid
            $profileHash.General.UuidPool = $sp.OperIdentPoolName
            $profileHash.General.Associated_Server = $sp.PnDn
            $profileHash.General.Template_Name = $templateName
            $profileHash.General.Template_Instance = $sp.OperSrcTemplName
            $profileHash.General.Assignment = @{}
            $pool = $sp | Get-UcsServerPoolAssignment
            if($pool.Count -gt 0) {
                $profileHash.General.Assignment.Server_Pool = $pool.Name
                $profileHash.General.Assignment.Qualifier = $pool.Qualifier
                $profileHash.General.Assignment.Restrict_Migration = $pool.RestrictMigration
            } else {
                $lsServer = $sp | Get-UcsLsBinding
                $profileHash.General.Assignment.Server = $lsServer.AssignedToDn
                $profileHash.General.Assignment.Restrict_Migration = $lsServer.RestrictMigration
            }

            # Service Profile Details - Storage Tab
            $profileHash.Storage = @{}
            $fcNode = $sp | Get-UcsVnicFcNode
            $vnicConn = $sp | Get-UcsVnicConnDef
            $profileHash.Storage.Nwwn = $fcNode.Addr
            $profileHash.Storage.Nwwn_Pool = $fcNode.IdentPoolName
            $profileHash.Storage.Local_Disk_Config = Get-UcsLocalDiskConfigPolicy -Dn $sp.OperLocalDiskPolicyName | Select-Object Mode,ProtectConfig,XtraProperty
            $profileHash.Storage.Connectivity_Policy = $vnicConn.SanConnPolicyName
            $profileHash.Storage.Connectivity_Instance = $vnicConn.OperSanConnPolicyName
            # Array variable for storing HBA configuration data
            $profileHash.Storage.Hbas = @()
            # Iterate through each HBA interface
            $sp | Get-UcsVhba | Sort-Object OperVcon,OperOrder | ForEach-Object {
                $hbaHash = @{}
                $hbaHash.Name = $_.Name
                $hbaHash.Pwwn = $_.Addr
                $hbaHash.FabricId = $_.SwitchId
                $hbaHash.Desired_Order = $_.Order
                $hbaHash.Actual_Order = $_.OperOrder
                $hbaHash.Desired_Placement = $_.AdminVcon
                $hbaHash.Actual_Placement = $_.OperVcon
                $hbaHash.EquipmentDn = $_.EquipmentDn
                $hbaHash.Vsan = ($_ | Get-UcsChild | Select-Object OperVnetName).OperVnetName
                $profileHash.Storage.Hbas += $hbaHash
            }

            # Service Profile Details - Network Tab
            $profileHash.Network = @{}
            $profileHash.Network.Connectivity_Policy = $vnicConn.LanConnPolicyName
            # Array variable for storing NIC configuration data
            $profileHash.Network.Nics = @()
            # Iterate through each vNIC and grab configuration data
            $sp | Get-UcsVnic | ForEach-Object {
                $nicHash = @{}
                $nicHash.Name = $_.Name
                $nicHash.Mac_Address = $_.Addr
                $nicHash.Desired_Order = $_.Order
                $nicHash.Actual_Order = $_.OperOrder
                $nicHash.Fabric_Id = $_.SwitchId
                $nicHash.Desired_Placement = $_.AdminVcon
                $nicHash.Actual_Placement = $_.OperVcon
                $nicHash.Mtu = $_.Mtu
                $nicHash.EquipmentDn = $_.EquipmentDn
                $nicHash.Adaptor_Profile = $_.AdaptorProfileName
                $nicHash.Control_Policy = $_.NwCtrlPolicyName
                $nicHash.Qos = $_.OperQosPolicyName
                $nicHash.Vlans = @()
                $nicHash.Vlans += $_ | Get-UcsChild -ClassId VnicEtherIf | Select-Object OperVnetName,Vnet,DefaultNet | Sort-Object {($_.Vnet) -as [int]}
                $profileHash.Network.Nics += $nicHash
            }

            # Service Profile Details - iSCSI vNICs
            $profileHash.iSCSI = @()
            # Iterate through all iSCSI interfaces and grab configuration data
            $sp | Get-UcsVnicIscsi | ForEach-Object {
                $iscsiHash = @{}
                $iscsiHash.Name = $_.Name
                $iscsiHash.Overlay = $_.VnicName
                $iscsiHash.Iqn = $_.InitiatorName
                $iscsiHash.Adapter_Policy = $_.AdaptorProfileName
                $iscsiHash.Mac = $_.Addr
                $iscsiHash.Vlan = ($_ | Get-UcsVnicVlan).VlanName
                $profileHash.iSCSI += $iscsiHash
            }

            # Service Profile Details - Performance
            $profileHash.Performance = @{}
            # Only grab performance data if the profile is associated
            if($profileHash.Assoc_State -eq 'associated') {
                # Get the collection time interval for adapter performance
                $interval = (Get-UcsCollectionPolicy -Name "adapter" | Select-Object CollectionInterval).CollectionInterval
                # Normalize collection interval to seconds
                Switch -wildcard (($interval -split '[0-9]')[-1]) {
                    "minute*" {$profileHash.Performance.Interval = ((($interval -split '[a-z]')[0]) -as [int]) * 60}
                    "second*" {$profileHash.Performance.Interval = ((($interval -split '[a-z]')[0]) -as [int])}
                }

                $cmd_args = @{
                    UcsStats = $statistics
                    RnFilter = "vnic-stats"
                    StatList = @("BytesRx","BytesRxDeltaAvg","BytesTx","BytesTxDeltaAvg","PacketsRx","PacketsRxDeltaAvg","PacketsTx","PacketsTxDeltaAvg")
                }
                # Iterate through each vHBA and grab performance data
                $profileHash.Performance.vHbas = @{}
                $profileHash.Storage.Hbas | ForEach-Object {
                    $profileHash.Performance.vHbas[$_.Name] = Get-DeviceStats @cmd_args -DnFilter $_.EquipmentDn
                }
                # Iterate through each vNIC and grab performance data
                $profileHash.Performance.vNics = @{}
                $profileHash.Network.Nics | ForEach-Object {
                    $profileHash.Performance.vNics[$_.Name] = Get-DeviceStats @cmd_args -DnFilter $_.EquipmentDn
                }
            }

            # Service Profile Policies
            $profileHash.Policies = @{}
            $profileHash.Policies.Bios = $sp.BiosProfileName
            $profileHash.Policies.Fw = $sp.HostFwPolicyName
            $profileHash.Policies.Ipmi = $sp.MgmtAccessPolicyName
            $profileHash.Policies.Power = $sp.PowerPolicyName
            $profileHash.Policies.Scrub = $sp.ScrubPolicyName
            $profileHash.Policies.Sol = $sp.SolPolicyName
            $profileHash.Policies.Stats = $sp.StatsPolicyName

            # Add current profile to template profile array
            $DomainHash.Profiles[$templateId].Profiles += $profileHash
        }
    }
    # End Service Profile Collection

    # Start LAN Configuration
    # Get the collection time interval for port performance
    $DomainHash.Collection = @{}
    $interval = (Get-UcsCollectionPolicy -Name "port" | Select-Object CollectionInterval).CollectionInterval
    # Normalize collection interval to seconds
    Switch -wildcard (($interval -split '[0-9]')[-1]) {
        "minute*" {$DomainHash.Collection.Port = ((($interval -split '[a-z]')[0]) -as [int]) * 60}
        "second*" {$DomainHash.Collection.Port = ((($interval -split '[a-z]')[0]) -as [int])}
    }
    # Uplink and Server Ports with Performance
    $DomainHash.Lan.UplinkPorts = @()
    $DomainHash.Lan.ServerPorts = @()
    # Iterate through each FI and collect port performance data based on port role
    $DomainHash.Inventory.FIs | ForEach-Object {
        # Uplink Ports
        $_.Ports | Where-Object IfRole -eq network | ForEach-Object {
            $port = $_
            $uplinkHash = @{}
            $uplinkHash.Dn = $_.Dn
            $uplinkHash.PortId = $_.PortId
            $uplinkHash.SlotId = $_.SlotId
            $uplinkHash.Fabric_Id = $_.SwitchId
            $uplinkHash.Mac = $_.Mac
            $uplinkHash.Speed = $_.OperSpeed
            $uplinkHash.IfType = $_.IfType
            $uplinkHash.IfRole = $_.IfRole
            $uplinkHash.XcvrType = $_.XcvrType
            $uplinkHash.Performance = @{}
            $cmd_args = @{
                UcsStats = $statistics
                DnFilter = "$($port.Dn)/.*stats"
                StatList = @("TotalBytes","TotalPackets","TotalBytesDeltaAvg")
            }
            $uplinkHash.Performance.Rx = Get-DeviceStats @cmd_args -RnFilter "rx[-]stats"
            $uplinkHash.Performance.Tx = Get-DeviceStats @cmd_args -RnFilter "tx[-]stats"
            $uplinkHash.Status = $_.OperState
            $uplinkHash.State = $_.AdminState
            $DomainHash.Lan.UplinkPorts += $uplinkHash
        }
        # Server Ports
        $_.Ports | Where-Object IfRole -eq server | ForEach-Object {
            $port = $_
            $serverPortHash = @{}
            $serverPortHash.Dn = $_.Dn
            $serverPortHash.PortId = $_.PortId
            $serverPortHash.SlotId = $_.SlotId
            $serverPortHash.Fabric_Id = $_.SwitchId
            $serverPortHash.Mac = $_.Mac
            $serverPortHash.Speed = $_.OperSpeed
            $serverPortHash.IfType = $_.IfType
            $serverPortHash.IfRole = $_.IfRole
            $serverPortHash.XcvrType = $_.XcvrType
            $serverPortHash.Performance = @{}
            $cmd_args = @{
                UcsStats = $statistics
                DnFilter = "$($port.Dn)/.*stats"
                StatList = @("TotalBytes","TotalPackets","TotalBytesDeltaAvg")
            }
            $serverPortHash.Performance.Rx = Get-DeviceStats @cmd_args -RnFilter "rx[-]stats"
            $serverPortHash.Performance.Tx = Get-DeviceStats @cmd_args -RnFilter "tx[-]stats"
            $serverPortHash.Status = $_.OperState
            $serverPortHash.State = $_.AdminState
            $DomainHash.Lan.ServerPorts += $serverPortHash
        }
    }
    # Fabric PortChannels
    $DomainHash.Lan.FabricPcs = @()
    Get-UcsFabricServerPortChannel -Ucs $handle | ForEach-Object {
        $uplinkHash = @{}
        $uplinkHash.Name = $_.Rn
        $uplinkHash.Chassis = $_.ChassisId
        $uplinkHash.Fabric_Id = $_.SwitchId
        $uplinkHash.Members = $_ | Get-UcsFabricServerPortChannelMember | Select-Object EpDn,PeerDn
        $DomainHash.Lan.FabricPcs += $uplinkHash
    }
    # Uplink PortChannels
    $DomainHash.Lan.UplinkPcs = @()
    Get-UcsUplinkPortChannel -Ucs $handle | ForEach-Object {
        $uplinkHash = @{}
        $uplinkHash.Name = $_.Rn
        $uplinkHash.Members = $_ | Get-UcsUplinkPortChannelMember | Select-Object EpDn,PeerDn
        $DomainHash.Lan.UplinkPcs += $uplinkHash
    }
    # Qos Domain Policies
    $DomainHash.Lan.Qos = @{}
    $DomainHash.Lan.Qos.Domain = @()
    $DomainHash.Lan.Qos.Domain += Get-UcsQosClass -Ucs $handle | Sort-Object Cos -Descending
    $DomainHash.Lan.Qos.Domain += Get-UcsBestEffortQosClass -Ucs $handle
    $DomainHash.Lan.Qos.Domain += Get-UcsFcQosClass -Ucs $handle

    # Qos Policies
    $DomainHash.Lan.Qos.Policies = @()
    Get-UcsQosPolicy -Ucs $handle | ForEach-Object {
        $qosHash = @{}
        $qosHash.Name = $_.Name
        $qosHash.Owner = $_.PolicyOwner
        ($qoshash.Burst,$qoshash.HostControl,$qoshash.Prio,$qoshash.Rate) = $_ | Get-UcsChild -ClassId EpqosEgress | Select-Object Burst,HostControl,Prio,Rate | ForEach-Object {$_.Burst,$_.HostControl,$_.Prio,$_.Rate}
        $DomainHash.Lan.Qos.Policies += $qosHash
    }

    # VLANs
    $DomainHash.Lan.Vlans = @()
    $DomainHash.Lan.Vlans += Get-UcsVlan -Ucs $handle | Where-Object {$_.IfRole -eq "network"} | Sort-Object -Property Ucs,Id

    # Network Control Policies
    $DomainHash.Lan.Control_Policies = @()
    $DomainHash.Lan.Control_Policies += Get-UcsNetworkControlPolicy -Ucs $handle | Where-Object Dn -ne "fabric/eth-estc/nwctrl-default" | Select-Object Cdp,MacRegisterMode,Name,UplinkFailAction,Descr,Dn,PolicyOwner

    # Mac Address Pool Definitions
    $DomainHash.Lan.Mac_Pools = @()
    Get-UcsMacPool -Ucs $handle | ForEach-Object {
        $macHash = @{}
        $macHash.Name = $_.Name
        $macHash.Assigned = $_.Assigned
        $macHash.Size = $_.Size
        ($macHash.From,$macHash.To) = $_ | Get-UcsMacMemberBlock | Select-Object From,To | ForEach-Object {$_.From,$_.To}
        $DomainHash.Lan.Mac_Pools += $macHash
    }

    # Mac Address Pool Allocations
    $DomainHash.Lan.Mac_Allocations = @()
    $DomainHash.Lan.Mac_Allocations += Get-UcsMacPoolPooled | Select-Object Id,Assigned,AssignedToDn

    # Ip Pool Definitions
    $DomainHash.Lan.Ip_Pools = @()
    Get-UcsIpPool -Ucs $handle | ForEach-Object {
        $ipHash = @{}
        $ipHash.Name = $_.Name
        $ipHash.Assigned = $_.Assigned
        $ipHash.Size = $_.Size
        ($ipHash.From,$ipHash.To,$ipHash.DefGw,$ipHash.Subnet,$ipHash.PrimDns) = $_ | Get-UcsIpPoolBlock | Select-Object From,To,DefGw,PrimDns,Subnet | ForEach-Object {$_.From,$_.To,$_.DefGw,$_.Subnet,$_.PrimDns}
        $DomainHash.Lan.Ip_Pools += $ipHash
    }

    # Ip Pool Allocations
    $DomainHash.Lan.Ip_Allocations = @()
    $DomainHash.Lan.Ip_Allocations += Get-UcsIpPoolPooled | Select-Object AssignedToDn,DefGw,Id,PrimDns,Subnet,Assigned

    # vNic Templates
    $DomainHash.Lan.vNic_Templates = @()
    $DomainHash.Lan.vNic_Templates += Get-UcsVnicTemplate -Ucs $handle | Select-Object Ucs,Dn,Name,Descr,SwitchId,TemplType,IdentPoolName,Mtu,NwCtrlPolicyName,QosPolicyName

    # End Lan Configuration

    # Start SAN Configuration
    # Uplink and Storage Ports
    $DomainHash.San.UplinkFcoePorts = @()
    $DomainHash.San.UplinkFcPorts = @()
    $DomainHash.San.StorageFcPorts = @()
    $DomainHash.San.StoragePorts = @()
    # Iterate through each FI and grab san performance data based on port role
    $DomainHash.Inventory.FIs | ForEach-Object {
        # SAN uplink ports
        $_.Ports | Where-Object IfRole -cmatch "fc.*uplink" | ForEach-Object {
            $port = $_
            $uplinkHash = @{}
            $uplinkHash.Dn = $_.Dn
            $uplinkHash.PortId = $_.PortId
            $uplinkHash.SlotId = $_.SlotId
            $uplinkHash.Fabric_Id = $_.SwitchId
            $uplinkHash.Mac = $_.Mac
            $uplinkHash.Speed = $_.OperSpeed
            $uplinkHash.IfType = $_.IfType
            $uplinkHash.IfRole = $_.IfRole
            $uplinkHash.XcvrType = $_.XcvrType
            $uplinkHash.Performance = @{}
            $cmd_args = @{
                UcsStats = $statistics
                DnFilter = "$($port.Dn)/.*stats"
                StatList = @("TotalBytes","TotalPackets","TotalBytesDeltaAvg")
            }
            $uplinkHash.Performance.Rx = Get-DeviceStats @cmd_args -RnFilter "rx[-]stats"
            $uplinkHash.Performance.Tx = Get-DeviceStats @cmd_args -RnFilter "tx[-]stats"
            $uplinkHash.Status = $_.OperState
            $uplinkHash.State = $_.AdminState
            $DomainHash.San.UplinkFcoePorts += $uplinkHash
        }
        # FC Uplink Ports
        $_.FcUplinkPorts | Where-Object IfRole -cmatch "network" | ForEach-Object {
            $port = $_
            $uplinkHash = @{}
            $uplinkHash.Dn = $_.Dn
            $uplinkHash.PortId = $_.PortId
            $uplinkHash.SlotId = $_.SlotId
            $uplinkHash.Fabric_Id = $_.SwitchId
            $uplinkHash.Wwn = $_.Wwn
            $uplinkHash.IfRole = $_.IfRole
            $uplinkHash.Speed = $_.OperSpeed
            $uplinkHash.Mode = $_.Mode
            $uplinkHash.XcvrType = $_.XcvrType
            $uplinkHash.Performance = @{}
            $cmd_args = @{
                UcsStats = $statistics
                DnFilter = "$($port.Dn)/stats"
                RnFilter = "stats"
                StatList = @("BytesRx","PacketsRx","BytesRxDeltaAvg","BytesTx","PacketsTx","BytesTxDeltaAvg")
            }
            $stats = Get-DeviceStats @cmd_args
            $uplinkHash.Performance.Rx = $stats | Select-Object BytesRx,PacketsRx,BytesRxDeltaAvg
            $uplinkHash.Performance.Tx = $stats | Select-Object BytesTx,PacketsTx,BytesTxDeltaAvg
            $uplinkHash.Status = $_.OperState
            $uplinkHash.State = $_.AdminState
            $DomainHash.San.UplinkFcPorts += $uplinkHash
        }
        # FC storage ports
        $_.FcUplinkPorts | Where-Object IfRole -cmatch "storage" | ForEach-Object {
            $port = $_
            $storageFcPortHash = @{}
            $storageFcPortHash.Dn = $_.Dn
            $storageFcPortHash.PortId = $_.PortId
            $storageFcPortHash.SlotId = $_.SlotId
            $storageFcPortHash.Fabric_Id = $_.SwitchId
            $storageFcPortHash.Wwn = $_.Wwn
            $storageFcPortHash.IfRole = $_.IfRole
            $storageFcPortHash.Speed = $_.OperSpeed
            $storageFcPortHash.Mode = $_.Mode
            $storageFcPortHash.XcvrType = $_.XcvrType
            $storageFcPortHash.Performance = @{}
            $cmd_args = @{
                UcsStats = $statistics
                DnFilter = "$($port.Dn)/stats"
                RnFilter = "stats"
                StatList = @("BytesRx","PacketsRx","BytesRxDeltaAvg","BytesTx","PacketsTx","BytesTxDeltaAvg")
            }
            $stats = Get-DeviceStats @cmd_args
            $storageFcPortHash.Performance.Rx = $stats | Select-Object BytesRx,PacketsRx,BytesRxDeltaAvg
            $storageFcPortHash.Performance.Tx = $stats | Select-Object BytesTx,PacketsTx,BytesTxDeltaAvg
            $storageFcPortHash.Status = $_.OperState
            $storageFcPortHash.State = $_.AdminState
            $DomainHash.San.StorageFcPorts += $storageFcPortHash
        }
        # Ethernet SAN storage ports
        $_.Ports | Where-Object IfRole -cmatch "storage" | ForEach-Object {
            $port = $_
            $storagePortHash = @{}
            $storagePortHash.Dn = $_.Dn
            $storagePortHash.PortId = $_.PortId
            $storagePortHash.SlotId = $_.SlotId
            $storagePortHash.Fabric_Id = $_.SwitchId
            $storagePortHash.Mac = $_.Mac
            $storagePortHash.Speed = $_.OperSpeed
            $storagePortHash.IfType = $_.IfType
            $storagePortHash.IfRole = $_.IfRole
            $storagePortHash.XcvrType = $_.XcvrType
            $storagePortHash.Performance = @{}
            $cmd_args = @{
                UcsStats = $statistics
                DnFilter = "$($port.Dn)/.*stats"
                StatList = @("TotalBytes","TotalPackets","TotalBytesDeltaAvg")
            }
            $storagePortHash.Performance.Rx = Get-DeviceStats @cmd_args -RnFilter "rx[-]stats"
            $storagePortHash.Performance.Tx = Get-DeviceStats @cmd_args -RnFilter "tx[-]stats"
            $storagePortHash.Status = $_.OperState
            $storagePortHash.State = $_.AdminState
            $DomainHash.San.StoragePorts += $storagePortHash
        }
    }
    # Uplink PortChannels
    $DomainHash.San.UplinkPcs = @()
    # Native FC PC uplinks
    Get-UcsFcUplinkPortChannel -Ucs $handle | ForEach-Object {
        $uplinkHash = @{}
        $uplinkHash.Name = $_.Rn
        $uplinkHash.Members = $_ | Get-UcsUplinkFcPort | Select-Object EpDn,PeerDn
        $DomainHash.San.UplinkPcs += $uplinkHash
    }
    # FCoE PC uplinks
    Get-UcsFabricFcoeSanPc -Ucs $handle | ForEach-Object {
        $uplinkHash = @{}
        $uplinkHash.Name = $_.Rn
        $uplinkHash.Members = $_ | Get-UcsFabricFcoeSanPcEp | Select-Object EpDn
        $DomainHash.San.FcoePcs += $uplinkHash
    }

    # VSANs
    $DomainHash.San.Vsans = @()
    $DomainHash.San.Vsans += Get-UcsVsan -Ucs $handle | Select-Object FcoeVlan,Id,name,SwitchId,ZoningState,IfRole,IfType,Transport

    # WWN Pools
    $DomainHash.San.Wwn_Pools = @()
    Get-UcsWwnPool -Ucs $handle | ForEach-Object {
        $wwnHash = @{}
        $wwnHash.Name = $_.Name
        $wwnHash.Assigned = $_.Assigned
        $wwnHash.Size = $_.Size
        $wwnHash.Purpose = $_.Purpose
        ($wwnHash.From,$wwnHash.To) = $_ | Get-UcsWwnMemberBlock | Select-Object From,To | ForEach-Object {$_.From,$_.To}
        $DomainHash.San.Wwn_Pools += $wwnHash
    }
    # WWN Allocations
    $DomainHash.San.Wwn_Allocations = @()
    $DomainHash.San.Wwn_Allocations += Get-UcsWwnInitiator | Select-Object AssignedToDn,Id,Assigned,Purpose

    # vHba Templates
    $DomainHash.San.vHba_Templates = Get-UcsVhbaTemplate -Ucs $handle | Select-Object Name,TempType

    # End San Configuration

    # Get Event List
    # Update current job progress
    $Process_Hash.Progress[$domain] = 84
    # Grab faults of critical, major, minor, and warning severity sorted by severity
    $faultList = Get-UcsFault -Ucs $handle -Filter 'Severity -cmatch "critical|major|minor|warning"' | Sort-Object -Property Ucs,Severity | Select-Object Ucs,Severity,Created,Descr,dn
    if($faultList) {
        # Iterate through each fault and grab information
        foreach ($fault in $faultList) {
            $faultHash = @{}
            $faultHash.Severity = $fault.Severity;
            $faultHash.Descr = $fault.Descr
            $faultHash.Dn = $fault.Dn
            $faultHash.Date = $fault.Created
            $DomainHash.Faults += $faultHash
        }
    }
    # Update current job progress
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


