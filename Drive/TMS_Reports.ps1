# TMS_Reports.ps1
# Description: Contains functions for generating various reports and analyses,
#              refactored to accept parameters for GUI use. Includes AAA Cooper integration.
#              Requires TMS_Helpers_General.ps1, carrier-specific helpers, and TMS_Config.ps1 to be loaded first.
#              This file should be dot-sourced by the main script.

# Assumes TMS_Helpers_General.ps1 functions are available.
# Assumes TMS_Config.ps1 variables are available.
# Assumes TMS_Carrier_*.ps1 functions (GUI versions) are available.

# --- Report Management (Console-based, GUI has its own interaction model) ---
function Manage-UserReports {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserReportsFolder
    )
    if (-not (Test-Path $UserReportsFolder)) { Write-Warning "No reports folder found at '$UserReportsFolder'."; Read-Host "Press Enter to continue..."; return }
    $exitReportMenu = $false
    while (-not $exitReportMenu) {
        if (Get-Command Clear-HostAndDrawHeader -ErrorAction SilentlyContinue) { Clear-HostAndDrawHeader -Title "Manage My Reports" -User (Split-Path $UserReportsFolder -Leaf) }
        else { Clear-Host; Write-Host "--- Manage My Reports (User: $(Split-Path $UserReportsFolder -Leaf)) ---"; Write-Warning "Clear-HostAndDrawHeader function not found." }
        
        $reportFiles = Get-ChildItem -Path $UserReportsFolder -Recurse -Filter "*.txt" -File | Sort-Object LastWriteTime -Descending
        Write-Host "`nAvailable Reports in '$UserReportsFolder':" -ForegroundColor Yellow
        if ($reportFiles.Count -gt 0) {
            for ($i = 0; $i -lt $reportFiles.Count; $i++) {
                $relativePath = $reportFiles[$i].FullName.Substring($UserReportsFolder.Length).TrimStart('\/')
                Write-Host (" [{0,2}] : {1} ({2:yyyy-MM-dd HH:mm})" -f ($i + 1), $relativePath, $reportFiles[$i].LastWriteTime)
            }
        }
        else { Write-Host "  No reports found." -ForegroundColor Gray }
        
        Write-Host "--------------------------------------" -ForegroundColor Blue
        Write-Host "Options:" -ForegroundColor Yellow
        Write-Host "  O. Open Report (Number)"
        Write-Host "  D. Delete Report (Number)"
        Write-Host "  X. Delete ALL Reports (Confirm)"
        Write-Host "  E. Open Reports Folder"
        Write-Host "  B. Back"
        Write-Host "--------------------------------------" -ForegroundColor Blue
        $reportChoice = Read-Host "Enter your choice"
        
        switch ($reportChoice.ToUpper()) {
            'O' {
                if ($reportFiles.Count -eq 0) { Write-Warning "No reports to open."; Read-Host "Press Enter..."; continue }
                $idxInput = Read-Host "Report number to open"
                if ($idxInput -match '^\d+$') {
                    $idx = [int]$idxInput - 1
                    if ($idx -ge 0 -and $idx -lt $reportFiles.Count) {
                        Write-Host "Opening '$($reportFiles[$idx].Name)'..."
                        if (Get-Command Open-FileExplorer -ErrorAction SilentlyContinue) {
                            Open-FileExplorer -Path $reportFiles[$idx].FullName
                        } else { Write-Warning "Open-FileExplorer function not found." }
                    } else { Write-Warning "Invalid report number." }
                } else { Write-Warning "Invalid input. Please enter a number." }
                Read-Host "`nPress Enter to continue..."
            }
            'D' {
                if ($reportFiles.Count -eq 0) { Write-Warning "No reports to delete."; Read-Host "Press Enter..."; continue }
                $idxInput = Read-Host "Report number to DELETE"
                if ($idxInput -match '^\d+$') {
                    $idx = [int]$idxInput - 1
                    if ($idx -ge 0 -and $idx -lt $reportFiles.Count) {
                        $fileToDelete = $reportFiles[$idx]
                        $confirm = Read-Host "Are you sure you want to DELETE '$($fileToDelete.Name)'? (Y/N)"
                        if ($confirm -match '^[Yy]$') {
                            try {
                                Remove-Item -Path $fileToDelete.FullName -Force -ErrorAction Stop
                                Write-Host "'$($fileToDelete.Name)' deleted successfully." -ForegroundColor Green
                            } catch { Write-Error "Failed to delete report: $($_.Exception.Message)" }
                        } else { Write-Host "Deletion cancelled." }
                    } else { Write-Warning "Invalid report number." }
                } else { Write-Warning "Invalid input. Please enter a number." }
                Read-Host "`nPress Enter to continue..."
            }
            'X' {
                if ($reportFiles.Count -eq 0) { Write-Warning "No reports to delete."; Read-Host "Press Enter..."; continue }
                $confirm = Read-Host "Are you sure you want to DELETE ALL $($reportFiles.Count) REPORTS? This cannot be undone. Type 'YES' to confirm."
                if ($confirm -eq 'YES') {
                    Write-Host "Deleting all reports..." -ForegroundColor Red
                    $deletedCount = 0; $failCount = 0
                    foreach ($file in $reportFiles) {
                        try {
                            Remove-Item -Path $file.FullName -Force -ErrorAction Stop
                            $deletedCount++
                        } catch {
                            Write-Error "Failed to delete '$($file.Name)': $($_.Exception.Message)"
                            $failCount++
                        }
                    }
                    Write-Host "Deletion complete. Deleted: ${deletedCount}. Failed: ${failCount}." -ForegroundColor Yellow
                } else { Write-Host "Deletion cancelled." }
                Read-Host "`nPress Enter to continue..."
            }
            'E' {
                if (Get-Command Open-FileExplorer -ErrorAction SilentlyContinue) {
                    Open-FileExplorer -Path $UserReportsFolder
                } else { Write-Warning "Open-FileExplorer function not found." }
                Read-Host "`nPress Enter to continue..."
            }
            'B' { $exitReportMenu = $true }
            Default { Write-Warning "Invalid choice."; Read-Host "Press Enter to continue..."; }
        }
    }
}


# --- Cross-Carrier Analysis Functions (Refactored for GUI) ---

function Run-CrossCarrierASPAnalysisGUI {
    # GUI VERSION: Calculates required margin for each carrier based on a desired ASP.
    # Optionally applies calculated margins (requires confirmation).
    param(
        [Parameter(Mandatory)][hashtable]$BrokerProfile,
        [Parameter(Mandatory)][hashtable]$SelectedCustomerProfile,
        [Parameter(Mandatory)][string]$ReportsBaseFolder, 
        [Parameter(Mandatory)][string]$UserReportsFolder, 
        [Parameter(Mandatory)][hashtable]$AllCentralKeys,
        [Parameter(Mandatory)][hashtable]$AllSAIAKeys,
        [Parameter(Mandatory)][hashtable]$AllRLKeys,
        [Parameter(Mandatory)][hashtable]$AllAverittKeys,
        [Parameter(Mandatory)][hashtable]$AllAAACooperKeys, 
        [Parameter(Mandatory)][decimal]$DesiredASPValue,
        [Parameter(Mandatory)][string]$CsvFilePath,
        [Parameter(Mandatory=$false)][bool]$ApplyMargins = $false,
        [Parameter(Mandatory=$false)][bool]$ASPFromHistory = $false
    )
    Write-Host "`nRunning Cross-Carrier ASP Analysis (GUI Mode)..." -ForegroundColor Cyan
    if ($DesiredASPValue -le 0) { Write-Error "Desired ASP must be positive."; return $null }

    $aspDisplayLabel = if ($ASPFromHistory) { "Avg Booked Price (History)" } else { "Desired Avg Selling Price" }
    Write-Host "Customer: '$($SelectedCustomerProfile.CustomerName)', ${aspDisplayLabel}: $($DesiredASPValue.ToString('C2')), CSV: '$CsvFilePath'"

    $analysisTitle = if ($ASPFromHistory) { "Margin Analysis Based on History ASP for $($SelectedCustomerProfile.CustomerName)" } else { "ASP Cross-Carrier Margin Analysis for $($SelectedCustomerProfile.CustomerName)" }
    $reportTypeSuffix = if ($ASPFromHistory) { "MarginByHistory" } else { "MarginForASP" }

    Write-Host "`nLoading and normalizing data for all carriers from '$CsvFilePath'..." -ForegroundColor Gray
    $requiredNormFuncs = @(
        "Load-And-Normalize-CentralData", "Load-And-Normalize-SAIAData",
        "Load-And-Normalize-RLData", "Load-And-Normalize-AverittData", "Load-And-Normalize-AAACooperData"
    )
    foreach ($funcName in $requiredNormFuncs) {
        if (-not (Get-Command $funcName -ErrorAction SilentlyContinue)) {
            Write-Error "Required data normalization function '$funcName' not found."
            return $null
        }
    }

    $centralShipmentData = Load-And-Normalize-CentralData -CsvPath $CsvFilePath
    $saiaShipmentData = Load-And-Normalize-SAIAData -CsvPath $CsvFilePath
    $rlShipmentData = Load-And-Normalize-RLData -CsvPath $CsvFilePath
    $averittShipmentData = Load-And-Normalize-AverittData -CsvPath $CsvFilePath
    $aaaCooperShipmentData = Load-And-Normalize-AAACooperData -CsvPath $CsvFilePath

    if (($null -eq $centralShipmentData -and $null -eq $saiaShipmentData -and $null -eq $rlShipmentData -and $null -eq $averittShipmentData -and $null -eq $aaaCooperShipmentData)) {
         Write-Error "Failed to load/normalize data for ALL carriers."
         return $null
    }
    
    $allShipmentDataSets = @($centralShipmentData, $saiaShipmentData, $rlShipmentData, $averittShipmentData, $aaaCooperShipmentData)
    $dataCounts = $allShipmentDataSets | ForEach-Object { if ($_) { $_.Count } else { 0 } } # Handle cases where a dataset might be null
    $totalRows = ($dataCounts | Measure-Object -Maximum).Maximum
    if ($totalRows -eq 0) { Write-Warning "CSV resulted in 0 processable rows after normalization for all carriers."; return $null }
    Write-Host " -> Data loaded and normalized (target ${totalRows} rows, individual counts may vary)." -ForegroundColor Green

    if (-not (Get-Command Get-PermittedKeys -ErrorAction SilentlyContinue)) { Write-Error "Get-PermittedKeys function not found."; return $null }
    $permittedCentral = Get-PermittedKeys -AllKeys $AllCentralKeys -AllowedKeyNames $SelectedCustomerProfile.AllowedCentralKeys
    $permittedSAIA = Get-PermittedKeys -AllKeys $AllSAIAKeys -AllowedKeyNames $SelectedCustomerProfile.AllowedSAIAKeys
    $permittedRL = Get-PermittedKeys -AllKeys $AllRLKeys -AllowedKeyNames $SelectedCustomerProfile.AllowedRLKeys
    $permittedAveritt = Get-PermittedKeys -AllKeys $AllAverittKeys -AllowedKeyNames $SelectedCustomerProfile.AllowedAverittKeys
    $permittedAAACooper = Get-PermittedKeys -AllKeys $AllAAACooperKeys -AllowedKeyNames $SelectedCustomerProfile.AllowedAAACooperKeys

    if ($permittedCentral.Count -eq 0 -and $permittedSAIA.Count -eq 0 -and $permittedRL.Count -eq 0 -and $permittedAveritt.Count -eq 0 -and $permittedAAACooper.Count -eq 0) {
        Write-Warning "No permitted keys/accounts found for Customer '$($SelectedCustomerProfile.CustomerName)'. Analysis cannot proceed."
        return $null
    }

    $aspResults = [System.Collections.Generic.List[object]]::new()
    $script:overallSkippedCount = 0 
    $requiredApiFuncs = @(
        "Invoke-CentralTransportApi", "Invoke-SAIAApi", "Invoke-RLApi",
        "Invoke-AverittApi", "Invoke-AAACooperApi"
    )
    foreach ($funcName in $requiredApiFuncs) {
        if (-not (Get-Command $funcName -ErrorAction SilentlyContinue)) {
            # Allow optional carriers to be missing their API functions
            if (($funcName -eq "Invoke-AverittApi" -and $AllAverittKeys.Count -eq 0) -or `
                ($funcName -eq "Invoke-AAACooperApi" -and $AllAAACooperKeys.Count -eq 0) ) {
                Write-Warning "API invocation function '$funcName' not found, but no keys are configured for this carrier. Skipping."
            } else {
                Write-Error "Required API invocation function '$funcName' not found."
                return $null
            }
        }
    }
    
    $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue';
    try {
        function Process-CarrierASP {
            param (
                [string]$CarrierName,
                [hashtable]$PermittedKeys,
                [array]$ShipmentDataset,
                [hashtable]$AllCarrierKeys, 
                [string]$CarrierKeyFolderPath 
            )
            if ($PermittedKeys.Count -gt 0 -and $ShipmentDataset -and $ShipmentDataset.Count -gt 0) { 
                Write-Host "`n--- Processing $CarrierName Accounts for $($SelectedCustomerProfile.CustomerName) ---" -ForegroundColor Yellow
                foreach ($keyName in ($PermittedKeys.Keys | Sort-Object)) {
                    $keyData = $PermittedKeys[$keyName]
                    Write-Host "Processing Account: ${keyName} (using $($ShipmentDataset.Count) ${CarrierName}-normalized rows)..." -ForegroundColor Cyan
                    $totalCostValue = 0.0; $processedCount = 0; $apiSkippedCount = 0
                    
                    foreach ($shipment in $ShipmentDataset) {
                        $costValue = $null
                        
                        # CORRECTED: Determine the correct function name, especially for Central Transport
                        $invokeFunctionName = if ($CarrierName -eq "Central") {
                                                "Invoke-CentralTransportApi"
                                            } elseif ($CarrierName -eq "AAACooper") { # Ensure exact match for any other specific names
                                                "Invoke-AAACooperApi"
                                            } else {
                                                "Invoke-${CarrierName}Api" 
                                            }

                        if (-not (Get-Command $invokeFunctionName -ErrorAction SilentlyContinue)) {
                            Write-Warning "API function '$invokeFunctionName' for carrier '$CarrierName' not found. Skipping this carrier for this shipment."
                            $apiSkippedCount++; continue
                        }
                        
                        $invokeParams = @{ KeyData = $keyData }
                        if ($CarrierName -eq "Central") { $invokeParams.ShipmentData = $shipment } 
                        elseif ($CarrierName -eq "SAIA") { $invokeParams.OriginZip = $shipment.OriginZip; $invokeParams.DestinationZip = $shipment.DestinationZip; $invokeParams.OriginCity = $shipment.OriginCity; $invokeParams.OriginState = $shipment.OriginState; $invokeParams.DestinationCity = $shipment.DestinationCity; $invokeParams.DestinationState = $shipment.DestinationState; $invokeParams.Details = $shipment.details }
                        elseif ($CarrierName -eq "RL") { $invokeParams.OriginZip = $shipment.OriginZip; $invokeParams.DestinationZip = $shipment.DestinationZip; $invokeParams.Commodities = $shipment.Commodities; $invokeParams.ShipmentDetails = $shipment }
                        elseif ($CarrierName -eq "Averitt") { $invokeParams.ShipmentData = $shipment } 
                        elseif ($CarrierName -eq "AAACooper") { $invokeParams.ShipmentData = $shipment } 
                        else { Write-Warning "Invoke function parameters not defined for $CarrierName"; $apiSkippedCount++; continue }

                        try {
                            $costValue = & $invokeFunctionName @invokeParams
                        } catch { Write-Warning "Error calling $invokeFunctionName for ${keyName}: $($_.Exception.Message)"; $costValue = $null }

                        if ($costValue -eq $null) { $apiSkippedCount++; continue }
                        $processedCount++; $totalCostValue += $costValue
                    } 

                    $avgCost = 0.0; $reqMargin = $null;
                    if ($processedCount -gt 0) {
                        $avgCost = $totalCostValue / $processedCount
                        if ($DesiredASPValue -ne 0) {
                            try { $reqMargin = (($DesiredASPValue - $avgCost) / $DesiredASPValue) * 100.0 }
                            catch { Write-Warning "Error calculating margin for ${keyName} ($CarrierName): $($_.Exception.Message)"}
                        }
                    }
                    $aspResults.Add([PSCustomObject]@{ Account = $keyName; Carrier = $CarrierName; AvgCost = $avgCost; RequiredMargin = $reqMargin; Processed = $processedCount; Skipped = $apiSkippedCount; KeysFolderPath = $CarrierKeyFolderPath; AllKeysRef = $AllCarrierKeys }); 
                    if ($processedCount -gt 0) { Write-Host " -> Avg Cost: $($avgCost.ToString('C2')). Required Margin: $(if($reqMargin -ne $null){'{0:N2}%' -f $reqMargin}else{'N/A'}) (${processedCount} processed, ${apiSkippedCount} API skips)" -ForegroundColor Green }
                    else { Write-Warning " -> No shipments processed successfully via API for $CarrierName account '${keyName}'. (${apiSkippedCount} API skips)"; }
                    $script:overallSkippedCount += $apiSkippedCount 
                } 
            } elseif ($PermittedKeys.Count -gt 0 -and ($null -eq $ShipmentDataset -or $ShipmentDataset.Count -eq 0)) { 
                Write-Warning "`nNo processable $CarrierName data rows from CSV for $($SelectedCustomerProfile.CustomerName), though permitted keys exist."
            } else { Write-Host "`nNo permitted $CarrierName accounts for $($SelectedCustomerProfile.CustomerName)." -ForegroundColor Gray }
        } 

        Process-CarrierASP -CarrierName "Central" -PermittedKeys $permittedCentral -ShipmentDataset $centralShipmentData -AllCarrierKeys $AllCentralKeys -CarrierKeyFolderPath $script:centralKeysFolderPath
        Process-CarrierASP -CarrierName "SAIA" -PermittedKeys $permittedSAIA -ShipmentDataset $saiaShipmentData -AllCarrierKeys $AllSAIAKeys -CarrierKeyFolderPath $script:saiaKeysFolderPath
        Process-CarrierASP -CarrierName "RL" -PermittedKeys $permittedRL -ShipmentDataset $rlShipmentData -AllCarrierKeys $AllRLKeys -CarrierKeyFolderPath $script:rlKeysFolderPath
        Process-CarrierASP -CarrierName "Averitt" -PermittedKeys $permittedAveritt -ShipmentDataset $averittShipmentData -AllCarrierKeys $AllAverittKeys -CarrierKeyFolderPath $script:averittKeysFolderPath
        Process-CarrierASP -CarrierName "AAACooper" -PermittedKeys $permittedAAACooper -ShipmentDataset $aaaCooperShipmentData -AllCarrierKeys $AllAAACooperKeys -CarrierKeyFolderPath $script:aaaCooperKeysFolderPath 

    } finally {
        $VerbosePreference = $CurrentVerbosePreference
    }

    $reportContent = [System.Collections.Generic.List[string]]::new()
    $reportContent.Add($analysisTitle); $reportContent.Add("=" * $analysisTitle.Length)
    $reportContent.Add("Date: $(Get-Date), Broker: $($BrokerProfile.Username), Input CSV: $(Split-Path $CsvFilePath -Leaf)")
    $reportContent.Add((" {0,-25}: {1:C2}" -f $aspDisplayLabel, $DesiredASPValue))
    $reportContent.Add("Total Unique CSV Rows Available (approx based on largest carrier set): ${totalRows}")
    $reportContent.Add("Total API Call Skips (sum across all accounts): ${script:overallSkippedCount}") 
    $reportContent.Add("")
    if ($aspResults.Count -gt 0) {
        $formattedTableOutput = ($aspResults | Sort-Object -Property Carrier, Account | Format-Table Carrier, Account, @{N='Avg Cost';E={if ($_.AvgCost -ne $null) {$_.AvgCost.ToString("C2")} else {'N/A'}}}, @{N='Required Margin %';E={if($_.RequiredMargin -ne $null){$_.RequiredMargin.ToString("N2") + "%"}else{"N/A"}}}, Processed, @{N='API Skips';E={$_.Skipped}} -AutoSize | Out-String)
        if ($null -ne $formattedTableOutput) {
            $linesToAdd = $formattedTableOutput.ToString().TrimEnd("`r", "`n").Split([Environment]::NewLine)
            foreach($lineToAdd in $linesToAdd){ $reportContent.Add($lineToAdd) }
        }
    } else { $reportContent.Add("No results generated (no permitted keys or no data processed).") }
    $reportContent.Add("-------------------------------------------------")

    if ($ApplyMargins -and $aspResults.Count -gt 0) {
        Write-Host "`nApplying calculated margins..." -ForegroundColor Yellow
        $reportContent.Add(""); $reportContent.Add("--- Margin Update Attempt Summary ---")
        $updateSuccessCount = 0; $updateFailCount = 0; $updateSkippedCount = 0; $updateInvalidMarginCount = 0;
        foreach ($result in $aspResults) {
            if ($result.Processed -eq 0) {
                $reportContent.Add("Skipping margin update for '$($result.Account)': No shipments were processed for cost calculation."); $updateSkippedCount++; continue
            }
            if ($result.RequiredMargin -eq $null -or -not ($result.RequiredMargin -is [double])) { $reportContent.Add("Skipping '$($result.Account)': Invalid required margin calculation."); $updateSkippedCount++; continue }
            
            $newMargin = [math]::Round($result.RequiredMargin, 1)
            if ($newMargin -lt 0 -or $newMargin -ge 100) { $reportContent.Add("Skipping '$($result.Account)': Calculated margin ${newMargin}% outside valid range (0-99.9)."); $updateInvalidMarginCount++; continue }

            if ([string]::IsNullOrWhiteSpace($result.KeysFolderPath) -or ($null -eq $result.AllKeysRef)) {
                Write-Warning "Keys folder path or AllKeys reference not found for '$($result.Account)' ($($result.Carrier)). Cannot update margin."
                $updateFailCount++; continue
            }
            if (-not (Get-Command Update-TariffMargin -ErrorAction SilentlyContinue)) { Write-Error "Update-TariffMargin function not found!"; $updateFailCount++; continue }

            Write-Host "Attempting to update margin for '$($result.Account)' (Carrier: $($result.Carrier)) to ${newMargin} % using folder '$($result.KeysFolderPath)'."
            $updateAttemptSuccess = Update-TariffMargin -TariffName $result.Account -AllKeysHashtable $result.AllKeysRef -KeysFolderPath $result.KeysFolderPath -NewMarginPercent $newMargin
            
            if ($updateAttemptSuccess) { $reportContent.Add("Success: '$($result.Account)' updated to ${newMargin}%."); $updateSuccessCount++ }
            else { $reportContent.Add("FAILED update for '$($result.Account)'. Check console/logs."); $updateFailCount++ }
        }
         $reportContent.Add("----------------------------------"); $reportContent.Add("Update Counts - Success: ${updateSuccessCount}, Failed: ${updateFailCount}, Skipped (No Data/Calc): ${updateSkippedCount}, Skipped (Invalid Range): ${updateInvalidMarginCount}")
    } elseif ($ApplyMargins) {
         $reportContent.Add(""); $reportContent.Add("NOTE: Apply Margins requested, but no results were generated to apply from.")
    }

    if (-not (Get-Command Get-ReportPath -ErrorAction SilentlyContinue)) { Write-Error "Get-ReportPath function not found."; return $null }
    $customerNameSafe = $SelectedCustomerProfile.CustomerName -replace '[^a-zA-Z0-9.-]', '_' 
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $BrokerProfile.Username -Carrier 'CrossCarrier' -ReportType $reportTypeSuffix -FilePrefix $customerNameSafe
    
    if ($reportFilePath) {
        try { $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -Force -ErrorAction Stop; Write-Host "`nReport saved: $reportFilePath" -ForegroundColor Green; return $reportFilePath }
        catch { Write-Error "Failed to save report: $($_.Exception.Message)"; return $null }
    } else { Write-Error "Failed to generate report path."; return $null }
}


function Run-MarginsByHistoryAnalysisGUI {
    param(
        [Parameter(Mandatory)][hashtable]$BrokerProfile,
        [Parameter(Mandatory)][hashtable]$SelectedCustomerProfile,
        [Parameter(Mandatory)][string]$ReportsBaseFolder,
        [Parameter(Mandatory)][string]$UserReportsFolder,
        [Parameter(Mandatory)][hashtable]$AllCentralKeys,
        [Parameter(Mandatory)][hashtable]$AllSAIAKeys,
        [Parameter(Mandatory)][hashtable]$AllRLKeys,
        [Parameter(Mandatory)][hashtable]$AllAverittKeys,
        [Parameter(Mandatory)][hashtable]$AllAAACooperKeys,
        [Parameter(Mandatory)][string]$CsvFilePath,
        [Parameter(Mandatory=$false)][bool]$ApplyMargins = $false
    )
    Write-Host "`nRunning Margins by History Analysis for Customer '$($SelectedCustomerProfile.CustomerName)' (GUI Mode)..." -ForegroundColor Cyan

    $totalBookedPrice = 0.0; $validBookedPriceCount = 0; $averageBookedPrice = $null
    try {
        if(-not (Test-Path $CsvFilePath -PathType Leaf)) { Write-Error "History CSV file not found: '$CsvFilePath'"; return $null }
        $rawData = Import-Csv -Path $CsvFilePath -ErrorAction Stop
        if ($rawData.Count -eq 0) { Write-Warning "History CSV '$CsvFilePath' is empty."; return $null }
        
        $header = $rawData[0].PSObject.Properties.Name
        $bookedPriceColumnName = $null
        $potentialPriceColumns = @('Booked Price', 'BookedPrice', 'Price', 'FinalQuotedPrice', 'Average Selling Price', 'CustomerRate', 'Sold For', 'Revenue')
        foreach($colName in $potentialPriceColumns){
            if($header -contains $colName){
                $bookedPriceColumnName = $colName
                break
            }
        }
        if($null -eq $bookedPriceColumnName){ Write-Error "History CSV '$CsvFilePath' missing a recognizable booked price column (e.g., 'Booked Price', 'Price', 'CustomerRate', 'Sold For')."; return $null }
        
        Write-Host "Calculating Average Booked Price from '$bookedPriceColumnName' column..." -ForegroundColor Gray
        foreach ($row in $rawData) {
            if ($row.$bookedPriceColumnName -ne $null -and $row.$bookedPriceColumnName -ne "") {
                try {
                    $priceString = $row.$bookedPriceColumnName -replace '[$,]' 
                    $bookedPrice = [decimal]$priceString
                    if ($bookedPrice -gt 0) { 
                        $totalBookedPrice += $bookedPrice
                        $validBookedPriceCount++
                    }
                } catch {
                    Write-Verbose "Skipping row due to conversion error for '${bookedPriceColumnName}': Value '$($row.$bookedPriceColumnName)'" 
                }
            }
        }
        if ($validBookedPriceCount -eq 0) { Write-Warning "No valid positive '${bookedPriceColumnName}' values found in '${CsvFilePath}' after attempting conversion."; return $null } 
        $averageBookedPrice = $totalBookedPrice / $validBookedPriceCount
        Write-Host " -> Average Booked Price: $($averageBookedPrice.ToString('C2')) (${validBookedPriceCount} valid rows)" -ForegroundColor Green 
    } catch { Write-Error "Error processing History CSV '${CsvFilePath}': $($_.Exception.Message)"; return $null } 

    if ($null -eq $averageBookedPrice) {
        Write-Error "Failed to calculate average booked price. Cannot proceed with CrossCarrierASPAnalysis."
        return $null
    }

    Write-Host "DEBUG: In Run-MarginsByHistoryAnalysisGUI, about to call Run-CrossCarrierASPAnalysisGUI." -ForegroundColor Magenta
    $TargetFunction = "Run-CrossCarrierASPAnalysisGUI"
    $CmdInfo = Get-Command $TargetFunction -ErrorAction SilentlyContinue
    if ($CmdInfo) {
        Write-Host "DEBUG: Found command '$TargetFunction'. Type: $($CmdInfo.CommandType)" -ForegroundColor Magenta
        if ($CmdInfo.Parameters) {
            Write-Host "DEBUG: Parameters for '$TargetFunction':" -ForegroundColor Magenta
            $CmdInfo.Parameters.Values | ForEach-Object { Write-Host "  - $($_.Name) (Type: $($_.ParameterType.FullName), Mandatory: $($_.Attributes | Where-Object {$_.TypeId.Name -eq 'ParameterAttribute'} | ForEach-Object {$_.Mandatory}))" -ForegroundColor Magenta }
            if ($CmdInfo.Parameters.ContainsKey("DesiredASPValue")) {
                Write-Host "DEBUG: Parameter 'DesiredASPValue' IS DEFINED for '$TargetFunction'." -ForegroundColor Green
            } else {
                Write-Host "DEBUG: Parameter 'DesiredASPValue' IS NOT DEFINED for '$TargetFunction'." -ForegroundColor Red
            }
        } else {
            Write-Host "DEBUG: No parameters found for '$TargetFunction'." -ForegroundColor Red
        }
    } else {
        Write-Host "DEBUG: Command '$TargetFunction' NOT FOUND." -ForegroundColor Red
    }
    Write-Host "DEBUG: Value of `$averageBookedPrice to be passed as DesiredASPValue: $averageBookedPrice (Type: $($averageBookedPrice.GetType().FullName))" -ForegroundColor Magenta
    Write-Host "DEBUG: Value of `$CsvFilePath to be passed: $CsvFilePath (Type: $($CsvFilePath.GetType().FullName))" -ForegroundColor Magenta

    if (-not (Get-Command Run-CrossCarrierASPAnalysisGUI -ErrorAction SilentlyContinue)) { Write-Error "Run-CrossCarrierASPAnalysisGUI function not found."; return $null }
    
    $crossCarrierParams = @{
        BrokerProfile             = $BrokerProfile
        SelectedCustomerProfile   = $SelectedCustomerProfile
        ReportsBaseFolder         = $ReportsBaseFolder
        UserReportsFolder         = $UserReportsFolder
        AllCentralKeys            = $AllCentralKeys
        AllSAIAKeys               = $AllSAIAKeys
        AllRLKeys                 = $AllRLKeys
        AllAverittKeys            = $AllAverittKeys
        AllAAACooperKeys          = $AllAAACooperKeys
        DesiredASPValue           = [decimal]$averageBookedPrice 
        CsvFilePath               = $CsvFilePath
        ApplyMargins              = $ApplyMargins
        ASPFromHistory            = $true
    }
    Write-Host "DEBUG: Calling Run-CrossCarrierASPAnalysisGUI with splatted parameters..." -ForegroundColor Yellow
    $reportPath = Run-CrossCarrierASPAnalysisGUI @crossCarrierParams

    return $reportPath 
}

Write-Verbose "TMS Reports Functions loaded (GUI Refactored with AAA Cooper Integration)."
