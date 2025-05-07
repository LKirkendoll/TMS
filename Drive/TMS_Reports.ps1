# TMS_Reports.ps1
# Description: Contains functions for cross-carrier reports, historical analysis,
#              and report management utilities.
#              Requires TMS_Helpers.ps1 and TMS_Config.ps1 to be loaded first (by main entry script).
#              This file should be dot-sourced by the main entry script (TMS_GUI.ps1).

# Assumes helper functions (Invoke-*, Load-And-Normalize-*, Get-ReportPath, etc.) are available.

function Run-CrossCarrierASPAnalysis {
    # Calculates the required margin for each permitted carrier/tariff for a customer
    # to meet a desired Average Selling Price (ASP).
    param(
        [Parameter(Mandatory=$true)][hashtable]$UserProfile, # The logged-in user's profile
        [Parameter(Mandatory=$true)][string]$ReportsBaseFolder, # e.g., $script:reportsBaseFolderPath
        [Parameter(Mandatory=$true)][hashtable]$AllCentralKeys,
        [Parameter(Mandatory=$true)][hashtable]$AllSAIAKeys,
        [Parameter(Mandatory=$true)][hashtable]$AllRLKeys,
        [Parameter(Mandatory=$true)][hashtable]$AllAverittKeys # <<< AVERITT ADDED >>>
    )

    Write-Host "`n--- Cross-Carrier ASP Margin Analysis ---" -ForegroundColor Yellow
    
    # 1. Select Customer Profile
    # If $Global:customerProfiles is not loaded or empty, attempt to load them.
    if ($null -eq $Global:customerProfiles -or $Global:customerProfiles.Count -eq 0) {
        Write-Warning "Customer profiles not loaded. Attempting to load..."
        if (Get-Command Load-AllCustomerProfiles -ErrorAction SilentlyContinue) {
            # $script:customerAccountsFolderPath should be set by the main GUI script
            $Global:customerProfiles = Load-AllCustomerProfiles -CustomerAccountsFolderPath $script:customerAccountsFolderPath
            if ($null -eq $Global:customerProfiles -or $Global:customerProfiles.Count -eq 0) {
                Write-Error "Failed to load customer profiles. Cannot proceed with Cross-Carrier ASP Analysis."
                return
            }
        } else {
            Write-Error "Load-AllCustomerProfiles function not found. Cannot proceed."
            return
        }
    }

    $customerNames = $Global:customerProfiles.Keys | Sort-Object
    if ($customerNames.Count -eq 0) { Write-Warning "No customer profiles loaded."; return }
    Write-Host "Select Customer for Analysis:"
    for ($i = 0; $i -lt $customerNames.Count; $i++) { Write-Host ("  [{0}] {1}" -f ($i + 1), $customerNames[$i]) }
    $custChoice = Read-Host "Enter number"
    if (-not ($custChoice -match '^\d+$' -and [int]$custChoice -ge 1 -and [int]$custChoice -le $customerNames.Count)) {
        Write-Warning "Invalid customer selection."; return
    }
    $selectedCustomerName = $customerNames[[int]$custChoice - 1]
    $customerProfile = $Global:customerProfiles[$selectedCustomerName]
    Write-Host "Analyzing for Customer: $($customerProfile.CustomerName)" -ForegroundColor Cyan

    # 2. Get Desired ASP
    $desiredASP = 0.0
    while ($desiredASP -le 0) {
        try { $desiredASP = [decimal](Read-Host "Enter Desired Average Selling Price (ASP) for $selectedCustomerName (e.g., 150.00)") } catch {}
        if ($desiredASP -le 0) { Write-Warning "ASP must be a positive number." }
    }

    # 3. Select CSV Data File
    # For cross-carrier, we need a generic CSV that can be adapted or a specific one if carriers have vastly different needs.
    # For now, assume a generic CSV that Load-And-Normalize-* functions can handle for their respective carriers.
    # Averitt reports will require its own specific detailed CSV.
    Write-Warning "Note: Each carrier's API may require specific data from the CSV."
    Write-Warning "For Averitt, ensure you select a CSV with detailed commodity and address information."
    $csvFilePath = Select-CsvFile -DialogTitle "Select Shipment Data CSV for Cross-Carrier Analysis" -InitialDirectory $script:shipmentDataFolderPath
    if (-not $csvFilePath) { Write-Warning "CSV file selection cancelled."; return }

    # --- Report Preparation ---
    $reportContent = [System.Collections.Generic.List[string]]::new()
    $userReportsFolder = Join-Path -Path $ReportsBaseFolder -ChildPath $UserProfile.Username
    $reportFilePath = Get-ReportPath -BaseDir $ReportsBaseFolder -Username $UserProfile.Username -Carrier 'CrossCarrier' -ReportType 'MarginForASP' -FilePrefix $selectedCustomerName
    if (-not $reportFilePath) { return }

    $reportContent.Add("Cross-Carrier Required Margin for Desired ASP Report")
    $reportContent.Add("User: $($UserProfile.Username)"); $reportContent.Add("Customer: $($customerProfile.CustomerName)"); $reportContent.Add("Date: $(Get-Date)")
    $reportContent.Add("Data File: $csvFilePath"); $reportContent.Add("Desired ASP: $($desiredASP.ToString("C2"))")
    $reportContent.Add(("-" * 80))
    $reportContent.Add(("Carrier".PadRight(20) + "Tariff/Account".PadRight(30) + "Avg Cost".PadRight(15) + "Req. Margin %".PadRight(15)))
    $reportContent.Add(("-" * 80))

    # --- Process Each Carrier ---
    $carriersToProcess = @(
        @{ Name = "Central"; AllKeys = $AllCentralKeys; NormalizeFunc = "Load-And-Normalize-CentralData"; InvokeFunc = "Invoke-CentralTransportApi"; AllowedKeysProp = "AllowedCentralKeys" },
        @{ Name = "SAIA"; AllKeys = $AllSAIAKeys; NormalizeFunc = "Load-And-Normalize-SAIAData"; InvokeFunc = "Invoke-SAIAApi"; AllowedKeysProp = "AllowedSAIAKeys" },
        @{ Name = "RL"; AllKeys = $AllRLKeys; NormalizeFunc = "Load-And-Normalize-RLData"; InvokeFunc = "Invoke-RLApi"; AllowedKeysProp = "AllowedRLKeys" },
        @{ Name = "Averitt"; AllKeys = $AllAverittKeys; NormalizeFunc = "Load-And-Normalize-AverittData"; InvokeFunc = "Invoke-AverittApi"; AllowedKeysProp = "AllowedAverittKeys" } # <<< AVERITT ADDED >>>
    )

    foreach ($carrierInfo in $carriersToProcess) {
        Write-Host "`nProcessing $($carrierInfo.Name)..." -ForegroundColor Green
        $permittedKeyNames = $customerProfile[$carrierInfo.AllowedKeysProp]
        if ($null -eq $permittedKeyNames -or $permittedKeyNames.Count -eq 0) {
            Write-Warning "No permitted keys for $($carrierInfo.Name) for customer $($customerProfile.CustomerName)."
            $reportContent.Add( ($carrierInfo.Name.PadRight(20) + "No Permitted Keys".PadRight(30) + "".PadRight(15) + "".PadRight(15)) )
            continue
        }
        $permittedKeys = Get-PermittedKeys -AllKeys $carrierInfo.AllKeys -AllowedKeyNames $permittedKeyNames
        if ($permittedKeys.Count -eq 0) {
            Write-Warning "Could not retrieve details for any permitted $($carrierInfo.Name) keys."
            $reportContent.Add( ($carrierInfo.Name.PadRight(20) + "Permitted Keys Not Found".PadRight(30) + "".PadRight(15) + "".PadRight(15)) )
            continue
        }

        # Load and normalize data ONCE per carrier type for this CSV
        # This assumes the CSV structure is compatible or adaptable by each NormalizeFunc
        $shipments = & $carrierInfo.NormalizeFunc -CsvPath $csvFilePath
        if ($null -eq $shipments -or $shipments.Count -eq 0) {
            Write-Warning "No processable shipments found in '$csvFilePath' for $($carrierInfo.Name) normalization."
            $reportContent.Add( ($carrierInfo.Name.PadRight(20) + "CSV Data Error".PadRight(30) + "".PadRight(15) + "".PadRight(15)) )
            continue
        }

        foreach ($keyName in ($permittedKeys.Keys | Sort-Object)) {
            $keyData = $permittedKeys[$keyName]
            $keyDisplayName = if ($keyData.ContainsKey('Name')) { $keyData.Name } else { $keyName }
            Write-Host "  Tariff: $keyDisplayName"

            $totalCost = 0.0; $processedCount = 0
            foreach ($shipmentRow in $shipments) {
                $cost = $null
                # --- API Call Logic ---
                # For Central, SAIA, RL, Invoke-Api function signature varies slightly.
                # Averitt's Invoke-AverittApi takes -KeyData and -ShipmentData (the whole row).
                # Other carriers might need specific parameters extracted.
                # This section needs careful adaptation if input requirements are very different.
                try {
                    $invokeParams = @{ KeyData = $keyData }
                    if ($carrierInfo.Name -eq "Central") {
                        $invokeParams.ShipmentData = $shipmentRow # Assumes $shipmentRow has 'Origin Postal Code', etc.
                    } elseif ($carrierInfo.Name -eq "SAIA") {
                        $invokeParams.OriginZip = $shipmentRow.OriginPostalCode; $invokeParams.DestinationZip = $shipmentRow.DestinationPostalCode
                        $invokeParams.OriginCity = $shipmentRow.OriginCity; $invokeParams.OriginState = $shipmentRow.OriginState
                        $invokeParams.DestinationCity = $shipmentRow.DestinationCity; $invokeParams.DestinationState = $shipmentRow.DestinationState
                        $invokeParams.Weight = $shipmentRow.'Total Weight'; $invokeParams.Class = $shipmentRow.'Freight Class 1'
                        $invokeParams.Details = $shipmentRow.details # From Load-And-Normalize-SAIAData
                    } elseif ($carrierInfo.Name -eq "RL") {
                        $invokeParams.OriginZip = $shipmentRow.OriginZip; $invokeParams.DestinationZip = $shipmentRow.DestinationZip
                        $invokeParams.Weight = $shipmentRow.Weight; $invokeParams.Class = $shipmentRow.Class
                        $invokeParams.ShipmentDetails = $shipmentRow # Pass the whole normalized row
                    } elseif ($carrierInfo.Name -eq "Averitt") {
                        $invokeParams.ShipmentData = $shipmentRow # Averitt takes the whole row
                    }
                    $cost = Invoke-Expression "$($carrierInfo.InvokeFunc) @invokeParams"
                } catch {
                    Write-Warning "Error invoking API for $($carrierInfo.Name) - ${keyDisplayName}: $($_.Exception.Message)"
                }
                # --- End API Call Logic ---

                if ($cost -ne $null -and $cost -is [decimal] -and $cost -gt 0) {
                    $totalCost += $cost
                    $processedCount++
                }
            } # End foreach shipmentRow

            if ($processedCount -gt 0) {
                $avgCost = $totalCost / $processedCount
                $reqMarginPercent = "N/A"
                if ($desiredASP -ne 0 -and $avgCost -ne 0) { # Avoid division by zero
                    try { $reqMarginPercent = (($desiredASP - $avgCost) / $desiredASP) * 100 } catch {}
                    if ($reqMarginPercent -is [double]) { $reqMarginPercent = $reqMarginPercent.ToString("N2") }
                }
                $reportContent.Add( ($carrierInfo.Name.PadRight(20) + $keyDisplayName.PadRight(30) + $avgCost.ToString("C").PadRight(15) + $reqMarginPercent.PadRight(15)) )
            } else {
                $reportContent.Add( ($carrierInfo.Name.PadRight(20) + $keyDisplayName.PadRight(30) + "No Rates".PadRight(15) + "N/A".PadRight(15)) )
                Write-Warning "No valid rates found for $($carrierInfo.Name) - $keyDisplayName to calculate average cost."
            }
        } # End foreach keyName
    } # End foreach carrierInfo

    $reportContent.Add(("-" * 80))
    $reportContent.Add("End of Report")

    # --- Save Report ---
    try {
        $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop
        Write-Host "`nCross-Carrier ASP Analysis Report saved to: $reportFilePath" -ForegroundColor Green
        if ((Read-Host "Open report file? (Y/N)").ToUpper() -eq 'Y') { Open-FileExplorer -Path $reportFilePath }
    } catch {
        Write-Error "Failed to save Cross-Carrier ASP Analysis report: $($_.Exception.Message)"
    }
}


function Manage-UserReports {
    # Helper function to open the current user's reports folder.
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserReportsFolder # Path to the specific user's reports folder
    )
    Clear-HostAndDrawHeader -Title "Manage My Reports" -User (Split-Path $UserReportsFolder -Leaf) # Username is the leaf of the path
    Write-Host "Your reports are stored in: $UserReportsFolder"
    if (Test-Path $UserReportsFolder -PathType Container) {
        $choice = Read-Host "Do you want to open this folder in File Explorer? (Y/N)"
        if ($choice -match '^[Yy]$') {
            Open-FileExplorer -Path $UserReportsFolder
        }
    } else {
        Write-Warning "Reports folder does not exist yet. It will be created when you run your first report."
    }
    Read-Host "`nPress Enter to return to the main menu..."
}

# Add other general or cross-carrier report functions here as needed.
# For example: Run-MarginsByHistoryAnalysis (would require significant logic for historical data processing)

Write-Verbose "TMS Reports Functions loaded."
