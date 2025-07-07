<#
Author: KZ 7/5/2025
#>

Get-NetAdapter | Select-Object Name, Status, InterfaceDescription, InterfaceGuid | Format-Table

Add-Type -AssemblyName Microsoft.VisualBasic

# Prompt for GUIDs
$input = [Microsoft.VisualBasic.Interaction]::InputBox("Enter one or more GUIDs to delete (comma or newline separated)", "GUID Cleanup")
if (-not $input) { Write-Host "No input provided. Exiting."; exit }

# Normalize input and ensure curly braces
$guids = $input -split '[,\n\r]+' | ForEach-Object {
    $g = $_.Trim()
    if ($g -and $g -notmatch '^\{.*\}$') {
        # Add curly braces if missing
        "{${g}}"
    } else {
        $g
    }
} | Where-Object { $_ -ne "" }

# Prompt for Test Mode
$testPrompt = [Microsoft.VisualBasic.Interaction]::MsgBox("Run in test mode (preview only)?", "YesNo,Question", "Test Mode")
$testMode = $testPrompt -eq "Yes"

<#
.SYNOPSIS
Removes class subkeys under the Network Adapter class that match the specified GUID.

.DESCRIPTION
This key contains network adapter configuration classes:
HKLM:\SYSTEM\ControlSet001\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}
Each subkey represents a network adapter instance, identified by the 'NetCfgInstanceId' property.
This function finds and deletes the subkey(s) matching the specified GUID.

.PARAMETER guid
The GUID string (with curly braces) representing the NetCfgInstanceId to remove.

.PARAMETER test
If true, previews deletions instead of performing them.
#>
function Remove-FromClassNetworkAdapters {
    param([string]$guid, [bool]$test)

    $classKey = "HKLM:\SYSTEM\ControlSet001\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}"

    if (-not (Test-Path $classKey)) {
        Write-Warning "Class key $classKey does not exist."
        return
    }

    if ($test) {
        Write-Host $classKey
    }

    # Arrays to hold subkeys for deletion or skipping
    $toDelete = @()
    $toSkip = @()

    Get-ChildItem -Path $classKey | ForEach-Object {
        $subkeyName = $_.PSChildName
        $subkeyPath = $_.PsPath
        try {
            $netCfgInstanceId = Get-ItemProperty -Path $subkeyPath -Name "NetCfgInstanceId" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty NetCfgInstanceId
            if ($netCfgInstanceId) {
                if ($netCfgInstanceId -ieq $guid) {
                    $toDelete += [PSCustomObject]@{
                        Subkey = $subkeyName
                        Path = $subkeyPath
                        NetCfgInstanceId = $netCfgInstanceId
                    }
                } else {
                    $toSkip += [PSCustomObject]@{
                        Subkey = $subkeyName
                        NetCfgInstanceId = $netCfgInstanceId
                    }
                }
            }
        } catch {
            Write-Warning "Error reading $subkeyPath : $_"
        }
    }

    # Process deletions
    foreach ($item in $toDelete) {
        if ($test) {
            Write-Host " ├─ Subkey: $($item.Subkey)"
            Write-Host " │    NetCfgInstanceId = $($item.NetCfgInstanceId)"
            Write-Host " [TEST] Would delete class subkey: $($item.Path)"
        } else {
            Remove-Item -Path $item.Path -Recurse -Force
            Write-Host "Deleted class subkey: $($item.Path) (NetCfgInstanceId = $($item.NetCfgInstanceId))"
        }
    }

    # Print skipped subkeys grouped once
    if ($test -and $toSkip.Count -gt 0) {
        Write-Host "`n [TEST] Skipping class subkeys:"
        foreach ($item in $toSkip) {
            Write-Host " ├─ Subkey: $($item.Subkey)"
            Write-Host " │    NetCfgInstanceId = $($item.NetCfgInstanceId)"
        }
    }
}

<#
.SYNOPSIS
Removes keys under the Control\Network branch matching the GUID.

.DESCRIPTION
The Control\Network branch stores network interface configuration and settings.
This function recursively searches under:
HKLM:\SYSTEM\ControlSet001\Control\Network
for subkeys named exactly as the GUID and deletes them.

.PARAMETER guid
The GUID string (with curly braces) to match key names.

.PARAMETER test
If true, previews deletions instead of performing them.
#>
function Remove-FromControlNetwork {
    param([string]$guid, [bool]$test)

    $baseKey = "HKLM:\SYSTEM\ControlSet001\Control\Network"

    if (-not (Test-Path $baseKey)) {
        Write-Warning "Key $baseKey does not exist."
        return
    }

    $foundKeys = Get-ChildItem -Path $baseKey -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -ieq $guid }

    if ($foundKeys.Count -eq 0) {
        if ($test) {
            Write-Host "[TEST] No matching keys found for $guid under Control\Network"
        }
        return
    }

    foreach ($key in $foundKeys) {
        if ($test) {
            Write-Host "[TEST] Would delete Control\Network key: $($key.PSPath)"
        } else {
            Remove-Item -Path $key.PSPath -Recurse -Force
            Write-Host "Deleted Control\Network key: $($key.PSPath)"
        }
    }
}

<#
.SYNOPSIS
Removes keys under the Control\NetworkSetup2 branch matching the GUID.

.DESCRIPTION
The Control\NetworkSetup2 branch holds network profile and setup information.
This function recursively searches under:
HKLM:\SYSTEM\ControlSet001\Control\NetworkSetup2
for subkeys named exactly as the GUID and deletes them.

.PARAMETER guid
The GUID string (with curly braces) to match key names.

.PARAMETER test
If true, previews deletions instead of performing them.
#>
function Remove-FromNetworkSetup2 {
    param([string]$guid, [bool]$test)

    $baseKey = "HKLM:\SYSTEM\ControlSet001\Control\NetworkSetup2"

    if (-not (Test-Path $baseKey)) {
        Write-Warning "Key $baseKey does not exist."
        return
    }

    $foundKeys = Get-ChildItem -Path $baseKey -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -ieq $guid }

    if ($foundKeys.Count -eq 0) {
        if ($test) {
            Write-Host "[TEST] No matching keys found for $guid under Control\NetworkSetup2"
        }
        return
    }

    foreach ($key in $foundKeys) {
        if ($test) {
            Write-Host "[TEST] Would delete NetworkSetup2 key: $($key.PSPath)"
        } else {
            Remove-Item -Path $key.PSPath -Recurse -Force
            Write-Host "Deleted NetworkSetup2 key: $($key.PSPath)"
        }
    }
}

<#
.SYNOPSIS
Deletes adapter subkeys matching the GUID under the L2Bridge service registry path.

.DESCRIPTION
The L2Bridge service is the Layer 2 Bridge service used for virtual switch networking,
including Hyper-V virtual switches.
This function targets:
HKLM:\SYSTEM\ControlSet001\Services\l2bridge\Parameters\Adapters\<GUID>
and removes the entire key for the specified GUID, if it exists.

.PARAMETER guid
The GUID string (with curly braces) identifying the adapter key to delete.

.PARAMETER test
If true, previews deletion instead of performing it.
#>
function Remove-FromL2BridgeAdapters {
    param([string]$guid, [bool]$test)

    $key = "HKLM:\SYSTEM\ControlSet001\Services\l2bridge\Parameters\Adapters"
    $adapterKey = Join-Path $key $guid

    if (Test-Path $adapterKey) {
        if ($test) {
            Write-Host "[TEST] Would delete registry key: $adapterKey"
        } else {
            Remove-Item -Path $adapterKey -Recurse -Force
            Write-Host "Deleted registry key: $adapterKey"
        }
    } elseif ($test) {
        Write-Host "[TEST] $adapterKey not found"
    }
}

<#
.SYNOPSIS
Removes specified GUID from the LanmanServer service linkage multi-string registry values.

.DESCRIPTION
The LanmanServer service provides file and print sharing over the network.
This function targets the registry path:
HKLM:\SYSTEM\ControlSet001\Services\LanmanServer\Linkage
and attempts to remove the given GUID from the multi-string values named 'Bind', 'Export', and 'Route'.
It updates those values by filtering out the GUID if present.

.PARAMETER guid
The GUID string (with curly braces) to remove from the linkage values.

.PARAMETER test
If true, the function only previews what it would remove without making changes.
#>
function Remove-FromLanManServer {
    param([string]$guid, [bool]$test)

    $keyPath = "HKLM:\SYSTEM\ControlSet001\Services\LanmanServer\Linkage"
    $subkeys = @("Bind", "Export", "Route")

    foreach ($subkey in $subkeys) {
        try {
            $values = Get-ItemProperty -Path $keyPath -Name $subkey -ErrorAction Stop
            $matches = $values.$subkey | Where-Object { $_ -match [regex]::Escape($guid) }

            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $subkey :`n$($matches -join "`n")"
                } else {
                    $updated = $values.$subkey | Where-Object { $_ -notmatch [regex]::Escape($guid) }
                    Set-ItemProperty -Path $keyPath -Name $subkey -Value $updated
                    Write-Host "Removed '$guid' from $subkey"
                }
            } elseif ($test) {
                Write-Host "[TEST] $guid not found in $subkey"
            }
        } catch {
            Write-Warning ("Could not process $subkey.`nReason: {0}" -f $_.Exception.Message)
        }
    }
}

<#
.SYNOPSIS
    Removes specified GUID entries from the LanmanWorkstation service Linkage registry keys.

.DESCRIPTION
    The LanmanWorkstation service handles SMB client functionality in Windows. 
    This function removes occurrences of a given GUID from the multi-string values 
    ("Bind", "Export", and "Route") in the registry key:
    HKLM:\SYSTEM\ControlSet001\Services\LanmanWorkstation\Linkage

.PARAMETER guid
    The GUID string to search for and remove.

.PARAMETER test
    If set to $true, runs in test mode and only displays what would be removed without making changes.
#>
function Remove-FromLanManWorkstation {
    param([string]$guid, [bool]$test)

    # Path to the LanmanWorkstation Linkage registry keys
    $keyPath = "HKLM:\SYSTEM\ControlSet001\Services\LanmanWorkstation\Linkage"
    $subkeys = @("Bind", "Export", "Route")

    foreach ($subkey in $subkeys) {
        $fullPath = "$keyPath\$subkey"
        try {
            $values = Get-ItemProperty -Path $keyPath -Name $subkey -ErrorAction Stop
            $matches = $values.$subkey | Where-Object { $_ -match $guid }

            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $subkey :`n$($matches -join "`n")"
                } else {
                    $updated = $values.$subkey | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $keyPath -Name $subkey -Value $updated
                    Write-Host "Removed '$guid' from $subkey"
                }
            } elseif ($test) {
                Write-Host "[TEST] $guid not found in $subkey"
            }
        } catch {
            Write-Warning ("Could not process $subkey.`nReason: {0}" -f $_.Exception.Message)
        }
    }
}

<#
.SYNOPSIS
    Removes specified GUID keys from NativeWifiP service adapter registry keys.

.DESCRIPTION
    This function deletes registry keys under 
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\NativeWifiP\Parameters\Adapters
    corresponding to the given GUIDs. The NativeWifiP service relates to 
    native Wi-Fi provisioning and management on Windows systems.

.PARAMETER guid
    The GUID string (with curly braces) representing the adapter key to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be deleted.

#>
function Remove-FromNativeWifiPAdapters {
    param([string]$guid, [bool]$test)

    $key = "HKLM:\SYSTEM\ControlSet001\Services\NativeWifiP\Parameters\Adapters"
    $adapterKey = Join-Path $key $guid

    if (Test-Path $adapterKey) {
        if ($test) {
            Write-Host "[TEST] Would delete registry key: $adapterKey"
        } else {
            Remove-Item -Path $adapterKey -Recurse -Force
            Write-Host "Deleted registry key: $adapterKey"
        }
    } elseif ($test) {
        Write-Host "[TEST] $adapterKey not found"
    }
}

<#
.SYNOPSIS
    Removes specified GUID keys from NdisCap service adapter registry keys.

.DESCRIPTION
    This function deletes registry keys under 
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\NdisCap\Parameters\Adapters
    corresponding to the given GUIDs. The NdisCap service is related to 
    Network Driver Interface Specification (NDIS) packet capture on Windows.

.PARAMETER guid
    The GUID string (with curly braces) representing the adapter key to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be deleted.

#>
function Remove-FromNdisCapAdapters {
    param([string]$guid, [bool]$test)

    $key = "HKLM:\SYSTEM\ControlSet001\Services\NdisCap\Parameters\Adapters"
    $adapterKey = Join-Path $key $guid

    if (Test-Path $adapterKey) {
        if ($test) {
            Write-Host "[TEST] Would delete registry key: $adapterKey"
        } else {
            Remove-Item -Path $adapterKey -Recurse -Force
            Write-Host "Deleted registry key: $adapterKey"
        }
    } elseif ($test) {
        Write-Host "[TEST] $adapterKey not found"
    }
}

<#
.SYNOPSIS
    Removes specified GUID entries from NetBIOS service Linkage keys.

.DESCRIPTION
    This function removes the given GUID(s) from the multi-string values 
    Bind, Export, and Route under 
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\NetBIOS\Linkage.
    The NetBIOS service provides legacy network basic input/output system 
    support, often for name resolution and session services in Windows networks.

.PARAMETER guid
    The GUID string (with curly braces) to remove from Linkage keys.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.

#>
function Remove-FromNetBIOS {
    param (
        [string]$guid,
        [bool]$test
    )

    $keyPath = "HKLM:\SYSTEM\ControlSet001\Services\NetBIOS\Linkage"
    $subkeys = @("Bind", "Export", "Route")

    foreach ($subkey in $subkeys) {
        $fullPath = "$keyPath\$subkey"
        try {
            $values = Get-ItemProperty -Path $keyPath -Name $subkey -ErrorAction Stop
            $matches = $values.$subkey | Where-Object { $_ -match $guid }

            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $subkey :`n$($matches -join "`n")"
                } else {
                    $updated = $values.$subkey | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $keyPath -Name $subkey -Value $updated
                    Write-Host "Removed '$guid' from $subkey"
                }
            }
            elseif ($test) {
                Write-Host "[TEST] $guid not found in $subkey"
            }
        }
        catch {
            Write-Warning ("Could not process $subkey.`nReason: {0}" -f $_.Exception.Message)
        }
    }
}

<#
.SYNOPSIS
    Removes specified GUID entries from NetBT service Linkage keys.

.DESCRIPTION
    This function removes the given GUID(s) from the multi-string values 
    Bind, Export, and Route under 
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\NetBT\Linkage.
    The NetBT (NetBIOS over TCP/IP) service provides NetBIOS services over 
    TCP/IP networks, enabling legacy name resolution and session services.

.PARAMETER guid
    The GUID string (with curly braces) to remove from Linkage keys.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.

#>
function Remove-FromNetBT {
    param (
        [string]$guid,
        [bool]$test
    )

    $keyPath = "HKLM:\SYSTEM\ControlSet001\Services\NetBT\Linkage"
    $subkeys = @("Bind", "Export", "Route")

    foreach ($subkey in $subkeys) {
        $fullPath = "$keyPath\$subkey"
        try {
            $values = Get-ItemProperty -Path $keyPath -Name $subkey -ErrorAction Stop
            $matches = $values.$subkey | Where-Object { $_ -match $guid }

            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $subkey :`n$($matches -join "`n")"
                } else {
                    $updated = $values.$subkey | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $keyPath -Name $subkey -Value $updated
                    Write-Host "Removed '$guid' from $subkey"
                }
            }
            elseif ($test) {
                Write-Host "[TEST] $guid not found in $subkey"
            }
        }
        catch {
            Write-Warning ("Could not process $subkey.`nReason: {0}" -f $_.Exception.Message)
        }
    }
}

<#
.SYNOPSIS
    Removes specified GUID entries from NetBT\Parameters\Interfaces.

.DESCRIPTION
    Deletes the entire Tcpip_{GUID} key from:
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\NetBT\Parameters\Interfaces
    which holds configuration for NetBT (NetBIOS over TCP/IP) interfaces.

.PARAMETER guid
    The GUID string (with curly braces) used to build the Tcpip_{GUID} subkey.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromNetBTInterfaces {
    param ([string]$guid, [bool]$test)

    $keyBase = "HKLM:\SYSTEM\ControlSet001\Services\NetBT\Parameters\Interfaces"
    $subkey = "Tcpip_$guid"
    $fullPath = Join-Path $keyBase $subkey

    if (Test-Path $fullPath) {
        if ($test) {
            Write-Host "[TEST] Would delete registry key: $fullPath"
        } else {
            Remove-Item -Path $fullPath -Recurse -Force
            Write-Host "Deleted registry key: $fullPath"
        }
    } elseif ($test) {
        Write-Host "[TEST] $fullPath not found"
    }
}

<#
.SYNOPSIS
    Removes GUID-based keys from nm3 service adapter entries.

.DESCRIPTION
    Deletes subkeys named by GUIDs from:
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\nm3\Parameters\Adapters
    which holds configuration for the Network Monitor 3 capture service.

.PARAMETER guid
    The GUID string (with curly braces) to delete.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromNm3Adapters {
    param ([string]$guid, [bool]$test)

    $baseKey = "HKLM:\SYSTEM\ControlSet001\Services\nm3\Parameters\Adapters"
    $fullPath = Join-Path $baseKey $guid

    if (Test-Path $fullPath) {
        if ($test) {
            Write-Host "[TEST] Would delete registry key: $fullPath"
        } else {
            Remove-Item -Path $fullPath -Recurse -Force
            Write-Host "Deleted registry key: $fullPath"
        }
    } elseif ($test) {
        Write-Host "[TEST] $fullPath not found"
    }
}

<#
.SYNOPSIS
    Removes GUID-based keys from npcap Parameters\Adapters.

.DESCRIPTION
    Deletes subkeys named by GUIDs from:
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\npcap\Parameters\Adapters
    which holds adapter configuration for Npcap packet capture drivers.

.PARAMETER guid
    The GUID string (with curly braces) to delete.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromNpcapAdapters {
    param ([string]$guid, [bool]$test)

    $baseKey = "HKLM:\SYSTEM\ControlSet001\Services\npcap\Parameters\Adapters"
    $fullPath = Join-Path $baseKey $guid

    if (Test-Path $fullPath) {
        if ($test) {
            Write-Host "[TEST] Would delete registry key: $fullPath"
        } else {
            Remove-Item -Path $fullPath -Recurse -Force
            Write-Host "Deleted registry key: $fullPath"
        }
    } elseif ($test) {
        Write-Host "[TEST] $fullPath not found"
    }
}

<#
.SYNOPSIS
    Removes GUID-based keys from npcap_wifi Parameters\Adapters.

.DESCRIPTION
    Deletes subkeys named by GUIDs from:
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\npcap_wifi\Parameters\Adapters
    which holds adapter configuration for WiFi interfaces used by Npcap.

.PARAMETER guid
    The GUID string (with curly braces) to delete.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromNpcapWifiAdapters {
    param ([string]$guid, [bool]$test)

    $baseKey = "HKLM:\SYSTEM\ControlSet001\Services\npcap_wifi\Parameters\Adapters"
    $fullPath = Join-Path $baseKey $guid

    if (Test-Path $fullPath) {
        if ($test) {
            Write-Host "[TEST] Would delete registry key: $fullPath"
        } else {
            Remove-Item -Path $fullPath -Recurse -Force
            Write-Host "Deleted registry key: $fullPath"
        }
    } elseif ($test) {
        Write-Host "[TEST] $fullPath not found"
    }
}

<#
.SYNOPSIS
    Removes GUID-based keys from Psched Parameters\Adapters.

.DESCRIPTION
    Deletes subkeys named by GUIDs from:
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\Psched\Parameters\Adapters
    which holds QoS (Quality of Service) policy scheduling settings for network interfaces.

.PARAMETER guid
    The GUID string (with curly braces) to delete.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromPschedAdapters {
    param ([string]$guid, [bool]$test)

    $baseKey = "HKLM:\SYSTEM\ControlSet001\Services\Psched\Parameters\Adapters"
    $fullPath = Join-Path $baseKey $guid

    if (Test-Path $fullPath) {
        if ($test) {
            Write-Host "[TEST] Would delete registry key: $fullPath"
        } else {
            Remove-Item -Path $fullPath -Recurse -Force
            Write-Host "Deleted registry key: $fullPath"
        }
    } elseif ($test) {
        Write-Host "[TEST] $fullPath not found"
    }
}

<#
.SYNOPSIS
    Removes GUID entries from RasPppoe service Linkage keys.

.DESCRIPTION
    This function removes the specified GUID from multi-string values
    Bind, Export, and Route under:
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\RasPppoe\Linkage.
    This service handles Point-to-Point Protocol over Ethernet (PPPoE).

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromRasPppoeLinkage {
    param ([string]$guid, [bool]$test)

    $keyPath = "HKLM:\SYSTEM\ControlSet001\Services\RasPppoe\Linkage"
    $subkeys = @("Bind", "Export", "Route")

    foreach ($subkey in $subkeys) {
        try {
            $values = Get-ItemProperty -Path $keyPath -Name $subkey -ErrorAction Stop
            $matches = $values.$subkey | Where-Object { $_ -match $guid }

            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $subkey :`n$($matches -join "`n")"
                } else {
                    $updated = $values.$subkey | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $keyPath -Name $subkey -Value $updated
                    Write-Host "Removed '$guid' from $subkey"
                }
            }
            elseif ($test) {
                Write-Host "[TEST] $guid not found in $subkey"
            }
        } catch {
            Write-Warning ("Could not process $subkey.`nReason: {0}" -f $_.Exception.Message)
        }
    }
}

<#
.SYNOPSIS
    Removes specified GUID from the Tcpip service linkage multi-string registry values.

.DESCRIPTION
    The Tcpip service handles TCP/IP protocol stack operations.
    This function targets the registry path:
    HKLM:\SYSTEM\ControlSet001\Services\Tcpip\Linkage
    and removes the given GUID from the multi-string values 'Bind', 'Export', and 'Route'.

.PARAMETER guid
    The GUID string (with curly braces) to remove from the linkage values.

.PARAMETER test
    If true, the function only previews what it would remove without making changes.
#>
function Remove-FromTcpipLinkage {
    param (
        [string]$guid,
        [bool]$test
    )

    $keyPath = "HKLM:\SYSTEM\ControlSet001\Services\Tcpip\Linkage"
    $subkeys = @("Bind", "Export", "Route")

    foreach ($subkey in $subkeys) {
        try {
            $values = Get-ItemProperty -Path $keyPath -Name $subkey -ErrorAction Stop
            $matches = $values.$subkey | Where-Object { $_ -match $guid }

            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $subkey :`n$($matches -join "`n")"
                } else {
                    $updated = $values.$subkey | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $keyPath -Name $subkey -Value $updated
                    Write-Host "Removed '$guid' from $subkey"
                }
            } elseif ($test) {
                Write-Host "[TEST] $guid not found in $subkey"
            }
        } catch {
            Write-Warning "Could not process $subkey. Reason: $_"
        }
    }
}

<#
.SYNOPSIS
    Deletes the GUID-named subkey from Tcpip Parameters Adapters.

.DESCRIPTION
    This function removes a registry subkey named with the GUID under:
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\Tcpip\Parameters\Adapters.
    These keys typically store adapter-specific TCP/IP settings.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromTcpipAdapters {
    param (
        [string]$guid,
        [bool]$test
    )

    $keyPath = "HKLM:\SYSTEM\ControlSet001\Services\Tcpip\Parameters\Adapters"
    $fullKey = Join-Path -Path $keyPath -ChildPath $guid

    if (Test-Path $fullKey) {
        if ($test) {
            Write-Host "[TEST] Would remove key: $fullKey"
        } else {
            Remove-Item -Path $fullKey -Recurse -Force
            Write-Host "Removed key: $fullKey"
        }
    } elseif ($test) {
        Write-Host "[TEST] $fullKey does not exist"
    }
}

<#
.SYNOPSIS
    Deletes the GUID-named subkey from Tcpip Parameters Interfaces.

.DESCRIPTION
    This function removes a registry subkey named with the GUID under:
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\Tcpip\Parameters\Interfaces.
    These keys store per-interface TCP/IP configuration.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromTcpipInterfaces {
    param (
        [string]$guid,
        [bool]$test
    )

    $keyPath = "HKLM:\SYSTEM\ControlSet001\Services\Tcpip\Parameters\Interfaces"
    $fullKey = Join-Path -Path $keyPath -ChildPath $guid

    if (Test-Path $fullKey) {
        if ($test) {
            Write-Host "[TEST] Would remove key: $fullKey"
        } else {
            Remove-Item -Path $fullKey -Recurse -Force
            Write-Host "Removed key: $fullKey"
        }
    } elseif ($test) {
        Write-Host "[TEST] $fullKey does not exist"
    }
}

<#
.SYNOPSIS
    Removes specified GUID from the Tcpip6 service linkage multi-string registry values.

.DESCRIPTION
    This function removes the given GUID from the multi-string values
    Bind, Export, and Route under:
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\Tcpip6\Linkage.
    The Tcpip6 service manages IPv6 stack functionality on Windows.

.PARAMETER guid
    The GUID string (with curly braces) to remove from linkage values.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromTcpip6Linkage {
    param (
        [string]$guid,
        [bool]$test
    )

    $keyPath = "HKLM:\SYSTEM\ControlSet001\Services\Tcpip6\Linkage"
    $subkeys = @("Bind", "Export", "Route")

    foreach ($subkey in $subkeys) {
        try {
            $values = Get-ItemProperty -Path $keyPath -Name $subkey -ErrorAction Stop
            $matches = $values.$subkey | Where-Object { $_ -match $guid }

            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $subkey :`n$($matches -join "`n")"
                } else {
                    $updated = $values.$subkey | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $keyPath -Name $subkey -Value $updated
                    Write-Host "Removed '$guid' from $subkey"
                }
            } elseif ($test) {
                Write-Host "[TEST] $guid not found in $subkey"
            }
        } catch {
            Write-Warning "Could not process $subkey. Reason: $_"
        }
    }
}

<#
.SYNOPSIS
    Removes GUID-named subkeys from Tcpip6 Parameters Interfaces.

.DESCRIPTION
    Deletes any subkeys matching the specified GUID under:
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\Tcpip6\Parameters\Interfaces.
    This key contains IPv6 interface parameters for TCP/IP stack.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromTcpip6Interfaces {
    param ([string]$guid, [bool]$test)

    $base = "HKLM:\SYSTEM\ControlSet001\Services\Tcpip6\Parameters\Interfaces"
    $target = Join-Path $base $guid

    if (Test-Path $target) {
        if ($test) {
            Write-Host "[TEST] Would delete: $target"
        } else {
            Remove-Item -Path $target -Recurse -Force
            Write-Host "Deleted: $target"
        }
    } elseif ($test) {
        Write-Host "[TEST] $target not found"
    }
}

<#
.SYNOPSIS
    Removes GUID entries from VMnetBridge service Linkage keys.

.DESCRIPTION
    Removes the specified GUID from multi-string values
    Bind, Export, and Route under:
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\VMnetBridge\Linkage.
    This service handles VMware virtual network bridging.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromVMnetBridge {
    param ([string]$guid, [bool]$test)

    $key = "HKLM:\SYSTEM\ControlSet001\Services\VMnetBridge\Linkage"
    $subkeys = @("Bind", "Export", "Route")

    foreach ($subkey in $subkeys) {
        try {
            $values = Get-ItemProperty -Path $key -Name $subkey -ErrorAction Stop
            $matches = $values.$subkey | Where-Object { $_ -match $guid }

            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $subkey in VMnetBridge"
                } else {
                    $updated = $values.$subkey | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $key -Name $subkey -Value $updated
                    Write-Host "Removed '$guid' from $subkey in VMnetBridge"
                }
            } elseif ($test) {
                Write-Host "[TEST] $guid not found in $subkey"
            }
        } catch {
            Write-Warning "Error processing VMnetBridge\$subkey : $_"
        }
    }
}

<#
.SYNOPSIS
    Removes GUID-matching subkeys from WFPLWFS Adapters.

.DESCRIPTION
    WFPLWFS is a lightweight filter platform driver. Deletes any subkeys matching the GUID under:
    HKEY_LOCAL_MACHINE\SYSTEM\ControlSet001\Services\WFPLWFS\Parameters\Adapters.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromWFPLWFSAdapters {
    param ([string]$guid, [bool]$test)

    $base = "HKLM:\SYSTEM\ControlSet001\Services\WFPLWFS\Parameters\Adapters"
    $target = Join-Path $base $guid

    if (Test-Path $target) {
        if ($test) {
            Write-Host "[TEST] Would delete: $target"
        } else {
            Remove-Item -Path $target -Recurse -Force
            Write-Host "Deleted: $target"
        }
    } elseif ($test) {
        Write-Host "[TEST] $target not found"
    }
}

<#
.SYNOPSIS
    Removes specified GUID entries from Control\Network key in CurrentControlSet.

.DESCRIPTION
    Removes the given GUID from multi-string values under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Network.
    This key contains network configuration entries.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentControlNetwork {
    param([string]$guid, [bool]$test)

    $keyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Network"

    try {
        $subkeys = Get-ChildItem -Path $keyPath -ErrorAction Stop | Where-Object { $_.PSIsContainer }

        foreach ($subkey in $subkeys) {
            $subkeyPath = $subkey.PSPath
            # For multi-string values or string values, you can customize if needed
            $props = Get-ItemProperty -Path $subkey.PSPath -ErrorAction SilentlyContinue
            if ($props) {
                foreach ($propName in $props.PSObject.Properties.Name) {
                    $value = $props.$propName
                    if ($value -is [System.Array]) {
                        # Multi-string
                        if ($value -contains $guid) {
                            if ($test) {
                                Write-Host "[TEST] Would remove GUID $guid from $propName in $subkeyPath"
                            } else {
                                $newValue = $value | Where-Object { $_ -ne $guid }
                                Set-ItemProperty -Path $subkey.PSPath -Name $propName -Value $newValue
                                Write-Host "Removed GUID $guid from $propName in $subkeyPath"
                            }
                        }
                    } elseif ($value -is [string]) {
                        if ($value -eq $guid) {
                            if ($test) {
                                Write-Host "[TEST] Would clear GUID $guid from $propName in $subkeyPath"
                            } else {
                                Set-ItemProperty -Path $subkey.PSPath -Name $propName -Value ""
                                Write-Host "Cleared GUID $guid from $propName in $subkeyPath"
                            }
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Failed to process CurrentControlSet Control\Network: $_"
    }
}

<#
.SYNOPSIS
    Removes specified GUID entries from Control\NetworkSetup2 key in CurrentControlSet.

.DESCRIPTION
    Removes the given GUID from multi-string values under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\NetworkSetup2.
    This key contains network setup parameters.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentNetworkSetup2 {
    param([string]$guid, [bool]$test)

    $keyPath = "HKLM:\SYSTEM\CurrentControlSet\Control\NetworkSetup2"

    try {
        $subkeys = Get-ChildItem -Path $keyPath -ErrorAction Stop | Where-Object { $_.PSIsContainer }

        foreach ($subkey in $subkeys) {
            $subkeyPath = $subkey.PSPath
            $props = Get-ItemProperty -Path $subkey.PSPath -ErrorAction SilentlyContinue
            if ($props) {
                foreach ($propName in $props.PSObject.Properties.Name) {
                    $value = $props.$propName
                    if ($value -is [System.Array]) {
                        if ($value -contains $guid) {
                            if ($test) {
                                Write-Host "[TEST] Would remove GUID $guid from $propName in $subkeyPath"
                            } else {
                                $newValue = $value | Where-Object { $_ -ne $guid }
                                Set-ItemProperty -Path $subkey.PSPath -Name $propName -Value $newValue
                                Write-Host "Removed GUID $guid from $propName in $subkeyPath"
                            }
                        }
                    } elseif ($value -is [string]) {
                        if ($value -eq $guid) {
                            if ($test) {
                                Write-Host "[TEST] Would clear GUID $guid from $propName in $subkeyPath"
                            } else {
                                Set-ItemProperty -Path $subkey.PSPath -Name $propName -Value ""
                                Write-Host "Cleared GUID $guid from $propName in $subkeyPath"
                            }
                        }
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Failed to process CurrentControlSet Control\NetworkSetup2: $_"
    }
}

<#
.SYNOPSIS
    Removes GUID-named keys under l2bridge Parameters Adapters in CurrentControlSet.

.DESCRIPTION
    Deletes subkeys named by GUID under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\l2bridge\Parameters\Adapters.
    This key holds adapters related to the Layer 2 Bridge service.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentL2BridgeAdapters {
    param([string]$guid, [bool]$test)

    $base = "HKLM:\SYSTEM\CurrentControlSet\Services\l2bridge\Parameters\Adapters"
    $target = Join-Path $base $guid

    if (Test-Path $target) {
        if ($test) {
            Write-Host "[TEST] Would delete registry key: $target"
        } else {
            Remove-Item -Path $target -Recurse -Force
            Write-Host "Deleted registry key: $target"
        }
    } elseif ($test) {
        Write-Host "[TEST] Registry key $target not found"
    }
}

<#
.SYNOPSIS
    Removes GUID-named keys under Control\Class network adapters in CurrentControlSet.

.DESCRIPTION
    Deletes subkeys named by GUID under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}.
    This key holds network adapter class driver settings.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentClassNetworkAdapters {
    param([string]$guid, [bool]$test)

    $base = "HKLM:\SYSTEM\CurrentControlSet\Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}"
    $subkeys = Get-ChildItem -Path $base -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer }

    $deleted = $false
    foreach ($subkey in $subkeys) {
        # The subkey names here are 4-digit numbers (e.g., 0000, 0001, 0002)
        # Each subkey can have a NetCfgInstanceId property which holds a GUID to match
        $netCfgInstanceId = (Get-ItemProperty -Path $subkey.PSPath -Name 'NetCfgInstanceId' -ErrorAction SilentlyContinue).NetCfgInstanceId
        if ($netCfgInstanceId -and $netCfgInstanceId -ieq $guid) {
            if ($test) {
                Write-Host "[TEST] Would delete network adapter key: $($subkey.PSPath) (NetCfgInstanceId = $netCfgInstanceId)"
            } else {
                Remove-Item -Path $subkey.PSPath -Recurse -Force
                Write-Host "Deleted network adapter key: $($subkey.PSPath) (NetCfgInstanceId = $netCfgInstanceId)"
            }
            $deleted = $true
        }
    }
    if (-not $deleted -and $test) {
        Write-Host "[TEST] No network adapter keys found with GUID $guid"
    }
}

<#
.SYNOPSIS
Removes specified GUID from the LanmanServer service linkage multi-string registry values.

.DESCRIPTION
The LanmanServer service provides file and print sharing over the network.
This function targets the registry path:
HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Linkage
and attempts to remove the given GUID from the multi-string values named 'Bind', 'Export', and 'Route'.
It updates those values by filtering out the GUID if present.

.PARAMETER guid
The GUID string (with curly braces) to remove from the linkage values.

.PARAMETER test
If true, the function only previews what it would remove without making changes.
#>
function Remove-FromCurrentLanManServer {
    param([string]$guid, [bool]$test)

    $keyPath = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Linkage"
    $subkeys = @("Bind", "Export", "Route")

    foreach ($subkey in $subkeys) {
        try {
            $values = Get-ItemProperty -Path $keyPath -Name $subkey -ErrorAction Stop
            $matches = $values.$subkey | Where-Object { $_ -match [regex]::Escape($guid) }

            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $subkey :`n$($matches -join "`n")"
                } else {
                    $updated = $values.$subkey | Where-Object { $_ -notmatch [regex]::Escape($guid) }
                    Set-ItemProperty -Path $keyPath -Name $subkey -Value $updated
                    Write-Host "Removed '$guid' from $subkey"
                }
            } elseif ($test) {
                Write-Host "[TEST] $guid not found in $subkey"
            }
        } catch {
            Write-Warning ("Could not process $subkey.`nReason: {0}" -f $_.Exception.Message)
        }
    }
}

<#
.SYNOPSIS
    Removes GUID entries from LanmanWorkstation service Linkage keys.

.DESCRIPTION
    Removes the specified GUID from multi-string values
    Bind, Export, and Route under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Linkage.
    This service handles the Workstation service for network connections.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentLanmanWorkstation {
    param ([string]$guid, [bool]$test)

    $key = "HKLM:\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Linkage"
    $subkeys = @("Bind", "Export", "Route")

    foreach ($subkey in $subkeys) {
        try {
            $values = Get-ItemProperty -Path $key -Name $subkey -ErrorAction Stop
            $matches = $values.$subkey | Where-Object { $_ -match $guid }

            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $subkey in LanmanWorkstation"
                } else {
                    $updated = $values.$subkey | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $key -Name $subkey -Value $updated
                    Write-Host "Removed '$guid' from $subkey in LanmanWorkstation"
                }
            } elseif ($test) {
                Write-Host "[TEST] $guid not found in $subkey"
            }
        } catch {
            Write-Warning "Error processing LanmanWorkstation\$subkey : $_"
        }
    }
}

<#
.SYNOPSIS
    Removes adapter GUID keys under NativeWifiP in CurrentControlSet.

.DESCRIPTION
    Deletes GUID-named subkeys under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NativeWifiP\Parameters\Adapters.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentNativeWifiPAdapters {
    param ([string]$guid, [bool]$test)

    $key = "HKLM:\SYSTEM\CurrentControlSet\Services\NativeWifiP\Parameters\Adapters"
    $target = Join-Path $key $guid

    if (Test-Path $target) {
        if ($test) {
            Write-Host "[TEST] Would delete: $target"
        } else {
            Remove-Item -Path $target -Recurse -Force
            Write-Host "Deleted: $target"
        }
    } elseif ($test) {
        Write-Host "[TEST] $target not found"
    }
}

<#
.SYNOPSIS
    Removes adapter GUID keys under NdisCap in CurrentControlSet.

.DESCRIPTION
    Deletes GUID-named subkeys under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NdisCap\Parameters\Adapters.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentNdisCapAdapters {
    param ([string]$guid, [bool]$test)

    $key = "HKLM:\SYSTEM\CurrentControlSet\Services\NdisCap\Parameters\Adapters"
    $target = Join-Path $key $guid

    if (Test-Path $target) {
        if ($test) {
            Write-Host "[TEST] Would delete: $target"
        } else {
            Remove-Item -Path $target -Recurse -Force
            Write-Host "Deleted: $target"
        }
    } elseif ($test) {
        Write-Host "[TEST] $target not found"
    }
}

<#
.SYNOPSIS
    Removes GUID entries from NetBIOS service Linkage keys.

.DESCRIPTION
    Removes the specified GUID from multi-string values
    Bind, Export, and Route under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NetBIOS\Linkage.
    This service provides legacy NetBIOS services over TCP/IP.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentNetBIOS {
    param ([string]$guid, [bool]$test)

    $key = "HKLM:\SYSTEM\CurrentControlSet\Services\NetBIOS\Linkage"
    $subkeys = @("Bind", "Export", "Route")

    foreach ($subkey in $subkeys) {
        try {
            $values = Get-ItemProperty -Path $key -Name $subkey -ErrorAction Stop
            $matches = $values.$subkey | Where-Object { $_ -match $guid }

            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $subkey in NetBIOS"
                } else {
                    $updated = $values.$subkey | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $key -Name $subkey -Value $updated
                    Write-Host "Removed '$guid' from $subkey in NetBIOS"
                }
            } elseif ($test) {
                Write-Host "[TEST] $guid not found in $subkey"
            }
        } catch {
            Write-Warning "Error processing NetBIOS\$subkey : $_"
        }
    }
}

<#
.SYNOPSIS
    Removes GUID entries from NetBT service Linkage keys.

.DESCRIPTION
    Removes the specified GUID from multi-string values
    Bind, Export, and Route under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NetBT\Linkage.
    This service provides NetBIOS over TCP/IP services.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentNetBT {
    param ([string]$guid, [bool]$test)

    $key = "HKLM:\SYSTEM\CurrentControlSet\Services\NetBT\Linkage"
    $subkeys = @("Bind", "Export", "Route")

    foreach ($subkey in $subkeys) {
        try {
            $values = Get-ItemProperty -Path $key -Name $subkey -ErrorAction Stop
            $matches = $values.$subkey | Where-Object { $_ -match $guid }

            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $subkey in NetBT"
                } else {
                    $updated = $values.$subkey | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $key -Name $subkey -Value $updated
                    Write-Host "Removed '$guid' from $subkey in NetBT"
                }
            } elseif ($test) {
                Write-Host "[TEST] $guid not found in $subkey"
            }
        } catch {
            Write-Warning "Error processing NetBT\$subkey : $_"
        }
    }
}

<#
.SYNOPSIS
    Removes GUID-named interface keys from NetBT Parameters Interfaces.

.DESCRIPTION
    Deletes subkeys named Tcpip_{GUID} under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NetBT\Parameters\Interfaces.
    This key contains network interface parameters for NetBIOS over TCP/IP.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentNetBTInterfaces {
    param ([string]$guid, [bool]$test)

    $base = "HKLM:\SYSTEM\CurrentControlSet\Services\NetBT\Parameters\Interfaces"
    $key = Join-Path $base ("Tcpip_" + $guid.Trim('{}'))

    if (Test-Path $key) {
        if ($test) {
            Write-Host "[TEST] Would delete: $key"
        } else {
            Remove-Item -Path $key -Recurse -Force
            Write-Host "Deleted: $key"
        }
    } elseif ($test) {
        Write-Host "[TEST] $key not found"
    }
}

<#
.SYNOPSIS
    Removes GUID-named adapter keys from nm3 Parameters Adapters.

.DESCRIPTION
    Deletes subkeys named after the GUID under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\nm3\Parameters\Adapters.
    Nm3 is a network filtering service used for certain VPN or network monitoring tools.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentNm3Adapters {
    param ([string]$guid, [bool]$test)

    $base = "HKLM:\SYSTEM\CurrentControlSet\Services\nm3\Parameters\Adapters"
    $key = Join-Path $base $guid

    if (Test-Path $key) {
        if ($test) {
            Write-Host "[TEST] Would delete: $key"
        } else {
            Remove-Item -Path $key -Recurse -Force
            Write-Host "Deleted: $key"
        }
    } elseif ($test) {
        Write-Host "[TEST] $key not found"
    }
}

<#
.SYNOPSIS
    Removes GUID-named adapter keys from npcap Parameters Adapters.

.DESCRIPTION
    Deletes subkeys named after the GUID under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\npcap\Parameters\Adapters.
    Npcap is a packet capture and network monitoring driver used by tools like Wireshark.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentNpcapAdapters {
    param ([string]$guid, [bool]$test)

    $base = "HKLM:\SYSTEM\CurrentControlSet\Services\npcap\Parameters\Adapters"
    $key = Join-Path $base $guid

    if (Test-Path $key) {
        if ($test) {
            Write-Host "[TEST] Would delete: $key"
        } else {
            Remove-Item -Path $key -Recurse -Force
            Write-Host "Deleted: $key"
        }
    } elseif ($test) {
        Write-Host "[TEST] $key not found"
    }
}

<#
.SYNOPSIS
    Removes GUID-named adapter keys from npcap_wifi Parameters Adapters.

.DESCRIPTION
    Deletes subkeys named after the GUID under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\npcap_wifi\Parameters\Adapters.
    This service handles wireless adapters for Npcap.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentNpcapWiFiAdapters {
    param ([string]$guid, [bool]$test)

    $base = "HKLM:\SYSTEM\CurrentControlSet\Services\npcap_wifi\Parameters\Adapters"
    $key = Join-Path $base $guid

    if (Test-Path $key) {
        if ($test) {
            Write-Host "[TEST] Would delete: $key"
        } else {
            Remove-Item -Path $key -Recurse -Force
            Write-Host "Deleted: $key"
        }
    } elseif ($test) {
        Write-Host "[TEST] $key not found"
    }
}

<#
.SYNOPSIS
    Removes GUID-named adapter keys from Psched Parameters Adapters.

.DESCRIPTION
    Deletes subkeys named after the GUID under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Psched\Parameters\Adapters.
    Psched is the QoS Packet Scheduler service managing network traffic priorities.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentPschedAdapters {
    param ([string]$guid, [bool]$test)

    $base = "HKLM:\SYSTEM\CurrentControlSet\Services\Psched\Parameters\Adapters"
    $key = Join-Path $base $guid

    if (Test-Path $key) {
        if ($test) {
            Write-Host "[TEST] Would delete: $key"
        } else {
            Remove-Item -Path $key -Recurse -Force
            Write-Host "Deleted: $key"
        }
    } elseif ($test) {
        Write-Host "[TEST] $key not found"
    }
}

<#
.SYNOPSIS
    Removes GUID entries from RasPppoe service Linkage keys.

.DESCRIPTION
    Removes the specified GUID from multi-string values Bind, Export, and Route under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\RasPppoe\Linkage.
    This service handles Point-to-Point Protocol over Ethernet (PPPoE).

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentRasPppoeLinkage {
    param ([string]$guid, [bool]$test)

    $key = "HKLM:\SYSTEM\CurrentControlSet\Services\RasPppoe\Linkage"
    foreach ($val in "Bind", "Export", "Route") {
        try {
            $cur = Get-ItemProperty -Path $key -Name $val -ErrorAction Stop
            $hits = $cur.$val | Where-Object { $_ -match $guid }
            if ($hits) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $val in RasPppoe"
                } else {
                    $new = $cur.$val | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $key -Name $val -Value $new
                    Write-Host "Removed '$guid' from $val in RasPppoe"
                }
            } elseif ($test) {
                Write-Host "[TEST] $guid not found in $val"
            }
        } catch {
            Write-Warning "Error in RasPppoe\$val : $_"
        }
    }
}

<#
.SYNOPSIS
    Removes GUID entries from Tcpip service Linkage keys.

.DESCRIPTION
    Removes the specified GUID from multi-string values Bind, Export, and Route under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Linkage.
    Tcpip is the core TCP/IP protocol driver service.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentTcpipLinkage {
    param ([string]$guid, [bool]$test)

    $key = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Linkage"
    foreach ($val in "Bind", "Export", "Route") {
        try {
            $cur = Get-ItemProperty -Path $key -Name $val -ErrorAction Stop
            $hits = $cur.$val | Where-Object { $_ -match $guid }
            if ($hits) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $val in Tcpip"
                } else {
                    $new = $cur.$val | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $key -Name $val -Value $new
                    Write-Host "Removed '$guid' from $val in Tcpip"
                }
            } elseif ($test) {
                Write-Host "[TEST] $guid not found in $val"
            }
        } catch {
            Write-Warning "Error in Tcpip\$val : $_"
        }
    }
}

<#
.SYNOPSIS
    Removes GUID-named adapter keys from Tcpip Parameters Adapters.

.DESCRIPTION
    Deletes subkeys named after the GUID under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Adapters.
    This key holds configuration data for TCP/IP adapters.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentTcpipAdapters {
    param ([string]$guid, [bool]$test)

    $base = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Adapters"
    $key = Join-Path $base $guid

    if (Test-Path $key) {
        if ($test) {
            Write-Host "[TEST] Would delete: $key"
        } else {
            Remove-Item -Path $key -Recurse -Force
            Write-Host "Deleted: $key"
        }
    } elseif ($test) {
        Write-Host "[TEST] $key not found"
    }
}

<#
.SYNOPSIS
    Removes GUID-named interface keys from Tcpip Parameters Interfaces.

.DESCRIPTION
    Deletes subkeys named after the GUID under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces.
    This key holds TCP/IP interface configurations.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentTcpipInterfaces {
    param ([string]$guid, [bool]$test)

    $base = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces"
    $key = Join-Path $base $guid

    if (Test-Path $key) {
        if ($test) {
            Write-Host "[TEST] Would delete: $key"
        } else {
            Remove-Item -Path $key -Recurse -Force
            Write-Host "Deleted: $key"
        }
    } elseif ($test) {
        Write-Host "[TEST] $key not found"
    }
}

<#
.SYNOPSIS
    Removes GUID entries from Tcpip6 service Linkage keys.

.DESCRIPTION
    Removes the specified GUID from multi-string values Bind, Export, and Route under:
    HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip6\Linkage.
    Tcpip6 is the TCP/IPv6 protocol driver service.

.PARAMETER guid
    The GUID string (with curly braces) to remove.

.PARAMETER test
    If true, runs in test mode and only shows what would be removed.
#>
function Remove-FromCurrentTcpip6Linkage {
    param ([string]$guid, [bool]$test)

    $key = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Linkage"
    foreach ($val in "Bind", "Export", "Route") {
        try {
            $cur = Get-ItemProperty -Path $key -Name $val -ErrorAction Stop
            $hits = $cur.$val | Where-Object { $_ -match $guid }
            if ($hits) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $val in Tcpip6"
                } else {
                    $new = $cur.$val | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $key -Name $val -Value $new
                    Write-Host "Removed '$guid' from $val in Tcpip6"
                }
            } elseif ($test) {
                Write-Host "[TEST] $guid not found in $val"
            }
        } catch {
            Write-Warning "Error in Tcpip6\$val : $_"
        }
    }
}

<#
.SYNOPSIS
Removes specified GUID-named interface keys from Tcpip6 Parameters Interfaces.

.DESCRIPTION
Deletes subkeys named with the GUID under:
HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\Interfaces.
This key contains IPv6 network interface parameters.

.PARAMETER guid
The GUID string (with curly braces) to remove.

.PARAMETER test
If true, the function only previews what it would delete without making changes.
#>
function Remove-FromCurrentTcpip6Interfaces {
    param (
        [string]$guid,
        [bool]$test
    )
    $basePath = "HKLM:\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\Interfaces"
    $targetKey = Join-Path $basePath $guid

    if (Test-Path $targetKey) {
        if ($test) {
            Write-Host "[TEST] Would delete key: $targetKey"
        } else {
            Remove-Item -Path $targetKey -Recurse -Force
            Write-Host "Deleted key: $targetKey"
        }
    } elseif ($test) {
        Write-Host "[TEST] Key not found: $targetKey"
    }
}

<#
.SYNOPSIS
Removes specified GUID from the VMnetBridge service linkage multi-string registry values.

.DESCRIPTION
The VMnetBridge service handles VMware bridged networking.
This function targets the registry path:
HKLM:\SYSTEM\CurrentControlSet\Services\VMnetBridge\Linkage
and attempts to remove the given GUID from the multi-string values named 'Bind', 'Export', and 'Route'.
It updates those values by filtering out the GUID if present.

.PARAMETER guid
The GUID string (with curly braces) to remove from the linkage values.

.PARAMETER test
If true, the function only previews what it would remove without making changes.
#>
function Remove-FromCurrentVMnetBridge {
    param (
        [string]$guid,
        [bool]$test
    )
    $linkagePath = "HKLM:\SYSTEM\CurrentControlSet\Services\VMnetBridge\Linkage"
    foreach ($valueName in @("Bind", "Export", "Route")) {
        try {
            $property = Get-ItemProperty -Path $linkagePath -Name $valueName -ErrorAction Stop
            $matches = $property.$valueName | Where-Object { $_ -match $guid }
            if ($matches) {
                if ($test) {
                    Write-Host "[TEST] Would remove '$guid' from $valueName in VMnetBridge"
                } else {
                    $updated = $property.$valueName | Where-Object { $_ -notmatch $guid }
                    Set-ItemProperty -Path $linkagePath -Name $valueName -Value $updated
                    Write-Host "Removed '$guid' from $valueName in VMnetBridge"
                }
            } elseif ($test) {
                Write-Host "[TEST] GUID not found in $valueName of VMnetBridge"
            }
        } catch {
            Write-Warning "Failed to process $valueName in VMnetBridge: $_"
        }
    }
}

<#
.SYNOPSIS
Removes GUID-named adapter keys from WFPLWFS Parameters Adapters.

.DESCRIPTION
The WFPLWFS service is a lightweight filter platform driver.
This function deletes subkeys named with the GUID under:
HKLM:\SYSTEM\CurrentControlSet\Services\WFPLWFS\Parameters\Adapters.

.PARAMETER guid
The GUID string (with curly braces) to remove.

.PARAMETER test
If true, the function only previews what it would delete without making changes.
#>
function Remove-FromCurrentWFPLWFSAdapters {
    param (
        [string]$guid,
        [bool]$test
    )
    $basePath = "HKLM:\SYSTEM\CurrentControlSet\Services\WFPLWFS\Parameters\Adapters"
    $targetKey = Join-Path $basePath $guid

    if (Test-Path $targetKey) {
        if ($test) {
            Write-Host "[TEST] Would delete key: $targetKey"
        } else {
            Remove-Item -Path $targetKey -Recurse -Force
            Write-Host "Deleted key: $targetKey"
        }
    } elseif ($test) {
        Write-Host "[TEST] Key not found: $targetKey"
    }
}

# Main loop: iterate over GUIDs and call all removal functions
foreach ($guid in $guids) {
    Write-Host "`nProcessing GUID: $guid"

    # Functions targeting ControlSet001 registry paths
    $controlSet001Functions = @(
        'Remove-FromClassNetworkAdapters',     # Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318}
        'Remove-FromControlNetwork',           # Control\Network
        'Remove-FromNetworkSetup2',            # Control\NetworkSetup2
        'Remove-FromL2BridgeAdapters',         # Services\l2bridge\Parameters\Adapters
        'Remove-FromLanManServer',             # Services\LanmanServer\Linkage
        'Remove-FromLanManWorkstation',        # Services\LanmanWorkstation\Linkage (ControlSet001)
        'Remove-FromNativeWifiPAdapters',      # Services\NativeWifiP\Parameters\Adapters (ControlSet001)
        'Remove-FromNdisCapAdapters',          # Services\NdisCap\Parameters\Adapters (ControlSet001)
        'Remove-FromNetBIOS',                  # Services\NetBIOS\Linkage (ControlSet001)
        'Remove-FromNetBT',                    # Services\NetBT\Linkage (ControlSet001)
        'Remove-FromNetBTInterfaces',          # Services\NetBT\Parameters\Interfaces (ControlSet001)
        'Remove-FromNm3Adapters',              # Services\nm3\Parameters\Adapters
        'Remove-FromNpcapAdapters',            # Services\npcap\Parameters\Adapters
        'Remove-FromNpcapWiFiAdapters',        # Services\npcap_wifi\Parameters\Adapters
        'Remove-FromPschedAdapters',           # Services\Psched\Parameters\Adapters
        'Remove-FromRasPppoeLinkage',          # Services\RasPppoe\Linkage
        'Remove-FromTcpipLinkage',             # Services\Tcpip\Linkage
        'Remove-FromTcpipAdapters',            # Services\Tcpip\Parameters\Adapters
        'Remove-FromTcpipInterfaces',          # Services\Tcpip\Parameters\Interfaces
        'Remove-FromTcpip6Linkage',            # Services\Tcpip6\Linkage
        'Remove-FromTcpip6Interfaces',         # Services\Tcpip6\Parameters\Interfaces (ControlSet001)
        'Remove-FromVMnetBridge',              # Services\VMnetBridge\Linkage
        'Remove-FromWFPLWFSAdapters'           # Services\WFPLWFS\Parameters\Adapters
    )

    # Functions targeting CurrentControlSet registry paths
    $currentControlSetFunctions = @(
        'Remove-FromCurrentControlNetwork',        # Control\Network (CurrentControlSet)
        'Remove-FromCurrentNetworkSetup2',         # Control\NetworkSetup2 (CurrentControlSet)
        'Remove-FromCurrentL2BridgeAdapters',      # Services\l2bridge\Parameters\Adapters (CurrentControlSet)
        'Remove-FromCurrentClassNetworkAdapters',  # Control\Class\{4d36e972-e325-11ce-bfc1-08002be10318} (CurrentControlSet)
        'Remove-FromCurrentLanmanServer',          # Services\LanmanServer\Linkage (CurrentControlSet)
        'Remove-FromCurrentLanmanWorkstation',     # Services\LanmanWorkstation\Linkage (CurrentControlSet)
        'Remove-FromCurrentNetBIOS',               # Services\NetBIOS\Linkage (CurrentControlSet)
        'Remove-FromCurrentNetBT',                 # Services\NetBT\Linkage (CurrentControlSet)
        'Remove-FromCurrentNetBTInterfaces',       # Services\NetBT\Parameters\Interfaces (CurrentControlSet)
        'Remove-FromCurrentNm3Adapters',           # Services\nm3\Parameters\Adapters
        'Remove-FromCurrentNativeWifiPAdapters',   # Services\NativeWifiP\Parameters\Adapters (CurrentControlSet)
        'Remove-FromCurrentNdisCapAdapters',       # Services\NdisCap\Parameters\Adapters (CurrentControlSet)
        'Remove-FromCurrentNpcapAdapters',         # Services\npcap\Parameters\Adapters (CurrentControlSet)
        'Remove-FromCurrentNpcapWiFiAdapters',     # Services\npcap_wifi\Parameters\Adapters (CurrentControlSet)
        'Remove-FromCurrentPschedAdapters',        # Services\Psched\Parameters\Adapters (CurrentControlSet)
        'Remove-FromCurrentRasPppoeLinkage',       # Services\RasPppoe\Linkage (CurrentControlSet)
        'Remove-FromCurrentTcpipLinkage',          # Services\Tcpip\Linkage (CurrentControlSet)
        'Remove-FromCurrentTcpipAdapters',         # Services\Tcpip\Parameters\Adapters (CurrentControlSet)
        'Remove-FromCurrentTcpipInterfaces',       # Services\Tcpip\Parameters\Interfaces (CurrentControlSet)
        'Remove-FromCurrentTcpip6Linkage',         # Services\Tcpip6\Linkage (CurrentControlSet)
        'Remove-FromCurrentTcpip6Interfaces',      # Services\Tcpip6\Parameters\Interfaces (CurrentControlSet)
        'Remove-FromCurrentVMnetBridge',           # Services\VMnetBridge\Linkage (CurrentControlSet)
        'Remove-FromCurrentWFPLWFSAdapters'        # Services\WFPLWFS\Parameters\Adapters (CurrentControlSet)
    )

    foreach ($func in $controlSet001Functions) {
        try {
            & $func -guid $guid -test:$testMode
        }
        catch {
            Write-Warning "Error executing $func for GUID $guid. Error: $_"
        }
    }

    foreach ($func in $currentControlSetFunctions) {
        try {
            & $func -guid $guid -test:$testMode
        }
        catch {
            Write-Warning "Error executing $func for GUID $guid. Error: $_"
        }
    }
}


