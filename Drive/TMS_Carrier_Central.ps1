# TMS_Carrier_Central.ps1
# Description: Contains functions specific to Central Transport operations,
#              refactored to accept parameters for GUI use.
#              Requires TMS_Helpers.ps1 and TMS_Config.ps1 to be loaded first (by main entry script).
#              This file should be dot-sourced by the main entry script (TMS_GUI.ps1).

# Assumes helper functions like Invoke-CentralTransportApi, Write-LoadingBar,
# Load-And-Normalize-CentralData, Get-ReportPath, Select-CsvFile,
# Select-SingleKeyEntry, Get-PermittedKeys are available from TMS_Helpers.ps1 or main script.
# Assumes config variables like $script:centralApiUri are available from TMS_Config.ps1 (via main script).

function Run-CentralComparisonReportGUI {
    # GUI VERSION: Generates a report comparing costs between two selected Central Transport keys/tariffs.
    param(
        [Parameter(Mandatory=$true)][hashtable]$Key1Data, # Pass the actual hashtable for Key 1
        [Parameter(Mandatory=$true)][hashtable]$Key2Data, # Pass the actual hashtable for Key 2
        [Parameter(Mandatory=$true)][string]$CsvFilePath,
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$UserReportsFolder # Base folder for the user
    )
    Write-Host "`nRunning Central Comparison Report (GUI Mode)..." -ForegroundColor Cyan
    # Ensure Name property exists for logging
    $key1DisplayName = if ($Key1Data.ContainsKey('Name')) { $Key1Data.Name } else { $Key1Data.TariffFileName | Split-Path -Leaf }
    $key2DisplayName = if ($Key2Data.ContainsKey('Name')) { $Key2Data.Name } else { $Key2Data.TariffFileName | Split-Path -Leaf }
    Write-Host "Comparing: '$key1DisplayName' vs '$key2DisplayName'"

    # --- Data Loading ---
    $shipments = Load-And-Normalize-CentralData -CsvPath $CsvFilePath
    if ($shipments -eq $null -or $shipments.Count -eq 0) {
        Write-Warning "No processable Central shipment data found in '$CsvFilePath'."
        return $null # Indicate failure
    }

    # --- Report Preparation ---
    $reportContent = [System.Collections.Generic.List[string]]::new()
    $resultsData = [System.Collections.Generic.List[object]]::new()
    $key1NameSafe = $key1DisplayName -replace '[^a-zA-Z0-9_-]', ''
    $key2NameSafe = $key2DisplayName -replace '[^a-zA-Z0-9_-]', ''
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $Username -Carrier 'Central' -ReportType 'Comparison' -FilePrefix ($key1NameSafe + "_vs_" + $key2NameSafe)
    if (-not $reportFilePath) { return $null } # Get-ReportPath will handle error messages

    $skippedShipmentCount = 0; $totalDifference = 0.0; $processedShipmentCount = 0
    $reportContent.Add("Central Transport Comparison Report"); $reportContent.Add("User: $Username"); $reportContent.Add("Date: $(Get-Date)"); $reportContent.Add("Data File: $CsvFilePath")
    $reportContent.Add("Comparing: '$key1DisplayName' vs '$key2DisplayName'")
    $reportContent.Add("----------------------------------------------------------------------")
    $col1WidthComp = 15; $col2WidthComp = 15; $col3WidthComp = 10; $col4WidthComp = 18; $col5WidthComp = 18; $col6WidthComp = 15 # Adjusted for potentially longer key names
    $headerLineComp = ("Origin Zip".PadRight($col1WidthComp)) + ("Dest Zip".PadRight($col2WidthComp)) + ("Weight".PadRight($col3WidthComp)) + ("$key1DisplayName Cost".PadRight($col4WidthComp)) + ("$key2DisplayName Cost".PadRight($col5WidthComp)) + ("Difference".PadRight($col6WidthComp))
    $reportContent.Add($headerLineComp); $reportContent.Add("----------------------------------------------------------------------")

    # --- Process Shipments ---
    $totalShipmentsToProcess = $shipments.Count
    Write-Host "Processing $totalShipmentsToProcess shipments for Central Comparison..."
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue # Check if available
    for ($i = 0; $i -lt $totalShipmentsToProcess; $i++) {
        $shipment = $shipments[$i]; $shipmentNumberForDisplay = $i + 1
        if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($shipmentNumberForDisplay * 100 / $totalShipmentsToProcess)) -Message "Processing Shipment $shipmentNumberForDisplay (Central Comp)" }

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue' # Suppress API call verbose for cleaner report run
        # Invoke-CentralTransportApi expects $ShipmentData (the normalized row) and $KeyData (the tariff details)
        $cost1 = Invoke-CentralTransportApi -ShipmentData $shipment -KeyData $Key1Data
        $cost2 = Invoke-CentralTransportApi -ShipmentData $shipment -KeyData $Key2Data
        $VerbosePreference = $CurrentVerbosePreference

        if ($cost1 -eq $null -or $cost2 -eq $null) {
            $skippedShipmentCount++; Write-Warning "Skipping Central shipment $shipmentNumberForDisplay (Origin: $($shipment.'Origin Postal Code')) due to API error or no rate."
            continue
        }

        $difference = $cost1 - $cost2; $totalDifference += $difference; $processedShipmentCount++
        $resultsData.Add([PSCustomObject]@{
            OriginZip = $shipment.'Origin Postal Code'
            DestZip = $shipment.'Destination Postal Code'
            Weight = $shipment.'Total Weight' # Already decimal from normalization
            Cost1 = $cost1
            Cost2 = $cost2
            Difference = $difference
        })
    }
    if ($useLoadingBar) { Write-Progress -Activity "Processing Central Comparison Shipments" -Completed }

    # --- Format Results ---
    foreach ($result in $resultsData) {
        try {
            $originZipStr = if ([string]::IsNullOrWhiteSpace($result.OriginZip)) { 'N/A' } else { $result.OriginZip }
            $destZipStr = if ([string]::IsNullOrWhiteSpace($result.DestZip)) { 'N/A' } else { $result.DestZip }
            $weightStr = if ($null -eq $result.Weight) { 'N/A' } else { $result.Weight.ToString("N0") }
            $cost1Str = if ($result.Cost1 -ne $null) { $result.Cost1.ToString("C2") } else { 'Error' }
            $cost2Str = if ($result.Cost2 -ne $null) { $result.Cost2.ToString("C2") } else { 'Error' }
            $diffStr = if ($result.Difference -ne $null) { $result.Difference.ToString("C2") } else { 'Error' }
            $line = ($originZipStr.PadRight($col1WidthComp)) + ($destZipStr.PadRight($col2WidthComp)) + ($weightStr.PadRight($col3WidthComp)) + ($cost1Str.PadRight($col4WidthComp)) + ($cost2Str.PadRight($col5WidthComp)) + ($diffStr.PadRight($col6WidthComp))
            $reportContent.Add($line)
        } catch {
            Write-Warning "Skipping result row (Origin: $($result.OriginZip)) due to formatting error: $($_.Exception.Message)"
            $skippedShipmentCount++; if($processedShipmentCount -gt 0){$processedShipmentCount--}; if($result.Difference -ne $null){$totalDifference -= $result.Difference}
        }
    }

    # --- Report Summary ---
    $reportContent.Add("----------------------------------------------------------------------")
    $reportContent.Add("Summary:")
    $reportContent.Add("Processed Shipments: $processedShipmentCount")
    $reportContent.Add("Skipped Shipments (API/Formatting Errors): $skippedShipmentCount")
    if ($processedShipmentCount -gt 0) {
        $avgDifference = if ($processedShipmentCount -ne 0) { $totalDifference / $processedShipmentCount } else { 0 }
        $reportContent.Add("Total Cost Difference ($key1DisplayName - $key2DisplayName): $($totalDifference.ToString("C2"))")
        $reportContent.Add("Average Cost Difference per Shipment: $($avgDifference.ToString("C2"))")
    } else { $reportContent.Add("No shipments could be processed successfully.") }
    $reportContent.Add("----------------------------------------------------------------------")
    $reportContent.Add("End of Report")

    # --- Save Report ---
    try {
        $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop
        Write-Host "`nCentral Comparison Report saved successfully to: $reportFilePath" -ForegroundColor Green
        return $reportFilePath # Return path on success
    } catch {
        Write-Error "Failed to save Central comparison report file '$reportFilePath': $($_.Exception.Message)"
        return $null # Return null on failure
    }
}


function Run-CentralMarginReportGUI {
     # GUI VERSION: Calculates the average margin required for a 'Comparison' key cost to match the average target sell price derived from a 'Base' key cost and its associated margin.
     param(
        [Parameter(Mandatory=$true)][hashtable]$BaseKeyData,      # Pass the actual hashtable for Base Key
        [Parameter(Mandatory=$true)][hashtable]$ComparisonKeyData, # Pass the actual hashtable for Comparison Key
        [Parameter(Mandatory=$true)][string]$CsvFilePath,
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$UserReportsFolder
     )
    Write-Host "`nRunning Central Average Required Margin Report (GUI Mode)..." -ForegroundColor Cyan
    $baseKeyName = if ($BaseKeyData.ContainsKey('Name')) { $BaseKeyData.Name } else { $BaseKeyData.TariffFileName | Split-Path -Leaf }
    $compKeyName = if ($ComparisonKeyData.ContainsKey('Name')) { $ComparisonKeyData.Name } else { $ComparisonKeyData.TariffFileName | Split-Path -Leaf }
    Write-Host "Base Cost Key: '$baseKeyName', Comparison Cost Key: '$compKeyName'"

    # --- Get Margin from Base Key File ---
    $customerCurrentMarginPercent = $null
    if ($BaseKeyData.ContainsKey('MarginPercent')) {
        try { $marginValue = [double]$BaseKeyData.MarginPercent; if ($marginValue -ge 0 -and $marginValue -lt 100) { $customerCurrentMarginPercent = $marginValue } else { Write-Warning "Margin '$($BaseKeyData.MarginPercent)' in '$baseKeyName' is invalid (must be 0-99.9)." } }
        catch { Write-Warning "Invalid MarginPercent value '$($BaseKeyData.MarginPercent)' in '$baseKeyName'." }
    } else { Write-Warning "'MarginPercent' not found in base key '$baseKeyName'." }

    if ($customerCurrentMarginPercent -eq $null) {
        if ($Global:PSBoundParameters.ContainsKey('DefaultMarginPercentage') -and $Global:DefaultMarginPercentage -ge 0 -and $Global:DefaultMarginPercentage -lt 100) {
            $customerCurrentMarginPercent = $Global:DefaultMarginPercentage
            Write-Warning "Using global default margin: $customerCurrentMarginPercent%"
        } else {
            Write-Error "Valid margin not found for '$baseKeyName' and no valid DefaultMarginPercentage configured globally. Cannot proceed."
            return $null
        }
    }
    $customerCurrentMarginDecimal = [decimal]$customerCurrentMarginPercent / 100.0
    Write-Host "Using Margin from Base Key '$baseKeyName': $customerCurrentMarginPercent%" -ForegroundColor Cyan

    # --- Data Loading ---
    $shipments = Load-And-Normalize-CentralData -CsvPath $CsvFilePath
    if ($shipments -eq $null -or $shipments.Count -eq 0) { Write-Warning "No processable Central shipment data found in '$CsvFilePath'."; return $null }

    # --- Report Preparation ---
    $reportContent = [System.Collections.Generic.List[string]]::new(); $resultsData = [System.Collections.Generic.List[object]]::new()
    $safeBaseKeyName = $baseKeyName -replace '[^a-zA-Z0-9_-]', ''; $safeCompKeyName = $compKeyName -replace '[^a-zA-Z0-9_-]', ''
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $Username -Carrier 'Central' -ReportType 'AvgMarginCalc' -FilePrefix ($safeBaseKeyName + "_vs_" + $safeCompKeyName)
    if (-not $reportFilePath) { return $null }

    $skippedShipmentCount = 0; $processedShipmentCount = 0; $totalTargetSellPrice = 0.0; $totalCompCost = 0.0
    $reportContent.Add("Central Transport Average Required Margin Calculation Report"); $reportContent.Add("User: $Username"); $reportContent.Add("Date: $(Get-Date)"); $reportContent.Add("Data File: $CsvFilePath")
    $reportContent.Add("Base Cost Key: '$baseKeyName'"); $reportContent.Add("Comparison Cost Key: '$compKeyName'"); $reportContent.Add("Margin Applied to Base Cost (from '$baseKeyName'): $($customerCurrentMarginPercent)%")
    $reportContent.Add("------------------------------------------------------------------------------------------------------------")
    $col1Width = 12; $col2Width = 12; $col3Width = 10; $col4Width = 15; $col5Width = 15; $col6Width = 18; $col7Width = 15; $col8Width = 15
    $headerLine = ("Origin Zip".PadRight($col1Width)) + ("Dest Zip".PadRight($col2Width)) + ("Weight".PadRight($col3Width)) + ("Base Cost".PadRight($col4Width)) + ("Comp Cost".PadRight($col5Width)) + ("Target Sell".PadRight($col6Width)) + ("Base Profit".PadRight($col7Width)) + ("Comp Profit".PadRight($col8Width))
    $reportContent.Add($headerLine); $reportContent.Add("------------------------------------------------------------------------------------------------------------")

    # --- Process Shipments ---
    $totalShipmentsToProcess = $shipments.Count; Write-Host "Processing $totalShipmentsToProcess shipments for Central Avg Margin Calc..."
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt $totalShipmentsToProcess; $i++) {
        $shipment = $shipments[$i]; $shipmentNumberForDisplay = $i + 1
        if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($shipmentNumberForDisplay * 100 / $totalShipmentsToProcess)) -Message "Processing Shipment $shipmentNumberForDisplay (Central Avg Margin)" }

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'
        $baseCost = Invoke-CentralTransportApi -ShipmentData $shipment -KeyData $BaseKeyData
        $compCost = Invoke-CentralTransportApi -ShipmentData $shipment -KeyData $ComparisonKeyData
        $VerbosePreference = $CurrentVerbosePreference

        if ($baseCost -eq $null -or $compCost -eq $null) { $skippedShipmentCount++; Write-Warning "Skipping Central shipment $shipmentNumberForDisplay (API error)."; continue }
        if ($baseCost -le 0) { $skippedShipmentCount++; Write-Warning "Skipping Central shipment $shipmentNumberForDisplay (zero/negative Base Cost)."; continue }

        $targetSellPrice = $null; $baseProfit = $null; $compProfit = $null
        try {
            if ((1.0 - $customerCurrentMarginDecimal) -eq 0) { throw "Division by zero (100% margin)" }
            $targetSellPrice = $baseCost / (1.0 - $customerCurrentMarginDecimal)
            $baseProfit = $targetSellPrice - $baseCost
            $compProfit = $targetSellPrice - $compCost
        } catch { Write-Warning "Skipping Central shipment $shipmentNumberForDisplay (calculation error: $($_.Exception.Message))"; $skippedShipmentCount++; continue }

        $processedShipmentCount++; $totalTargetSellPrice += $targetSellPrice; $totalCompCost += $compCost
        $resultsData.Add([PSCustomObject]@{
            OriginZip = $shipment.'Origin Postal Code'; DestZip = $shipment.'Destination Postal Code'; Weight = $shipment.'Total Weight'
            BaseCost = $baseCost; CompCost = $compCost; TargetSellPrice = $targetSellPrice; BaseProfit = $baseProfit; CompProfit = $compProfit
        })
    }
    if ($useLoadingBar) { Write-Progress -Activity "Processing Central Avg Margin Shipments" -Completed }

    # --- Format Results ---
    foreach ($result in $resultsData) {
        try {
             $originZipStr = if ([string]::IsNullOrWhiteSpace($result.OriginZip)) { 'N/A' } else { $result.OriginZip }
             $destZipStr = if ([string]::IsNullOrWhiteSpace($result.DestZip)) { 'N/A' } else { $result.DestZip }
             $weightStr = if ($null -eq $result.Weight) { 'N/A' } else { $result.Weight.ToString("N0") }
             $baseCostStr = if ($result.BaseCost -ne $null) { $result.BaseCost.ToString("C2") } else { 'Error' }
             $compCostStr = if ($result.CompCost -ne $null) { $result.CompCost.ToString("C2") } else { 'Error' }
             $targetSellPriceStr = if ($result.TargetSellPrice -ne $null) { $result.TargetSellPrice.ToString("C2") } else { 'Error' }
             $baseProfitStr = if ($result.BaseProfit -ne $null) { $result.BaseProfit.ToString("C2") } else { 'Error' }
             $compProfitStr = if ($result.CompProfit -ne $null) { $result.CompProfit.ToString("C2") } else { 'Error' }
             $line = ($originZipStr.PadRight($col1Width)) + ($destZipStr.PadRight($col2Width)) + ($weightStr.PadRight($col3Width)) + ($baseCostStr.PadRight($col4Width)) + ($compCostStr.PadRight($col5Width)) + ($targetSellPriceStr.PadRight($col6Width)) + ($baseProfitStr.PadRight($col7Width)) + ($compProfitStr.PadRight($col8Width))
             $reportContent.Add($line)
        } catch { Write-Warning "Skipping result row (Origin: $($result.OriginZip)) due to formatting error: $($_.Exception.Message)"; $skippedShipmentCount++ }
    }

    # --- Report Summary ---
    $reportContent.Add("------------------------------------------------------------------------------------------------------------")
    $reportContent.Add("Summary:")
    $reportContent.Add("Processed Shipments (API & Calc OK): $processedShipmentCount")
    $reportContent.Add("Skipped Shipments (API/Calc/Formatting Errors): $skippedShipmentCount")
    if ($processedShipmentCount -gt 0) {
        $avgTargetSellPrice = $totalTargetSellPrice / $processedShipmentCount
        $avgCompCost = $totalCompCost / $processedShipmentCount
        $reportContent.Add("Average Target Sell Price (from '$baseKeyName' + $customerCurrentMarginPercent%): $($avgTargetSellPrice.ToString("C2"))")
        $reportContent.Add("Average Comparison Cost (from '$compKeyName'): $($avgCompCost.ToString("C2"))")
        $avgRequiredMarginPercent = $null
        if ($avgTargetSellPrice -ne 0) {
            try {
                $avgRequiredMarginDecimal = ($avgTargetSellPrice - $avgCompCost) / $avgTargetSellPrice
                $avgRequiredMarginPercent = $avgRequiredMarginDecimal * 100.0
                $reportContent.Add("Average Margin Required on '$compKeyName' Cost to Match Average Target Sell Price: $($avgRequiredMarginPercent.ToString("N2"))%")
            } catch { $reportContent.Add("Could not calculate average required margin due to calculation error.")}
        } else { $reportContent.Add("Average Target Sell Price is zero, cannot calculate average required margin.") }
    } else { $reportContent.Add("No shipments processed successfully, cannot calculate average margin.") }
    $reportContent.Add("------------------------------------------------------------------------------------------------------------")
    $reportContent.Add("End of Report")

    # --- Save Report ---
    try { $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop; Write-Host "`nCentral Average Margin Calculation Report saved successfully to: $reportFilePath" -ForegroundColor Green; return $reportFilePath }
    catch { Write-Error "Failed to save Central report file '$reportFilePath': $($_.Exception.Message)"; return $null }
}


function Calculate-CentralMarginForASPReportGUI {
    # GUI VERSION: Calculates the required margin for a specific key/tariff to meet a desired ASP.
    param(
        [Parameter(Mandatory=$true)][hashtable]$CostAccountInfo, # Pass the actual hashtable for the selected key
        [Parameter(Mandatory=$true)][decimal]$DesiredASP,      # Pass the desired ASP value
        [Parameter(Mandatory=$true)][string]$CsvFilePath,
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$UserReportsFolder
    )
    Write-Host "`nRunning Central Required Margin for Desired ASP Report (GUI Mode)..." -ForegroundColor Cyan
    if ($DesiredASP -le 0) { Write-Error "Desired ASP must be positive."; return $null }
    $costAccountName = if ($CostAccountInfo.ContainsKey('Name')) { $CostAccountInfo.Name } else { $CostAccountInfo.TariffFileName | Split-Path -Leaf }
    Write-Host "Cost Basis Account: '$costAccountName', Desired ASP: $($DesiredASP.ToString('C2'))"

    # --- Data Loading ---
    $shipments = Load-And-Normalize-CentralData -CsvPath $CsvFilePath
    if ($shipments -eq $null -or $shipments.Count -eq 0) { Write-Warning "No processable Central shipment data found in '$CsvFilePath'."; return $null }

    # --- Report Preparation ---
    $reportContent = [System.Collections.Generic.List[string]]::new(); $resultsData = [System.Collections.Generic.List[object]]::new()
    $safeCostKeyName = $costAccountName -replace '[^a-zA-Z0-9_-]', ''
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $Username -Carrier 'Central' -ReportType 'MarginForASP' -FilePrefix $safeCostKeyName
    if (-not $reportFilePath) { return $null }

    $skippedShipmentCount = 0; $processedShipmentCount = 0; $totalCostValue = 0.0
    $reportContent.Add("Central Transport Required Margin for Desired ASP Report"); $reportContent.Add("User: $Username"); $reportContent.Add("Date: $(Get-Date)"); $reportContent.Add("Data File: $CsvFilePath")
    $reportContent.Add("Cost Basis Account: '$costAccountName'"); $reportContent.Add("Desired Average Sell Price (ASP): $($DesiredASP.ToString("C2"))")
    $reportContent.Add("--------------------------------------------------------------------------")
    $col1ASP = 12; $col2ASP = 12; $col3ASP = 15; $col4ASP = 15
    $headerLineASP = ("Origin Zip".PadRight($col1ASP)) + ("Dest Zip".PadRight($col2ASP)) + ("Weight".PadRight($col3ASP)) + ("Retrieved Cost".PadRight($col4ASP))
    $reportContent.Add($headerLineASP); $reportContent.Add("--------------------------------------------------------------------------")

    # --- Process Shipments ---
    $totalShipmentsToProcess = $shipments.Count; Write-Host "Processing $totalShipmentsToProcess shipments for Central Margin for ASP..."
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt $totalShipmentsToProcess; $i++) {
        $shipment = $shipments[$i]; $shipmentNumberForDisplay = $i + 1
        if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($shipmentNumberForDisplay * 100 / $totalShipmentsToProcess)) -Message "Processing Shipment $shipmentNumberForDisplay (Central Margin/ASP)" }

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'
        $costValue = Invoke-CentralTransportApi -ShipmentData $shipment -KeyData $CostAccountInfo
        $VerbosePreference = $CurrentVerbosePreference

        if ($costValue -eq $null) { $skippedShipmentCount++; Write-Warning "Skipping Central shipment $shipmentNumberForDisplay (API error)."; continue }

        $processedShipmentCount++; $totalCostValue += $costValue
        $resultsData.Add([PSCustomObject]@{
            OriginZip = $shipment.'Origin Postal Code'; DestZip = $shipment.'Destination Postal Code'; Weight = $shipment.'Total Weight'; Cost = $costValue
        })
    }
    if ($useLoadingBar) { Write-Progress -Activity "Processing Central Margin/ASP Shipments" -Completed }

    # --- Format Results ---
    foreach ($result in $resultsData) {
        try {
             $originZipStr = if ([string]::IsNullOrWhiteSpace($result.OriginZip)) { 'N/A' } else { $result.OriginZip }
             $destZipStr = if ([string]::IsNullOrWhiteSpace($result.DestZip)) { 'N/A' } else { $result.DestZip }
             $weightStr = if ($null -eq $result.Weight) { 'N/A' } else { $result.Weight.ToString("N0") }
             $costStr = if ($result.Cost -ne $null) { $result.Cost.ToString("C2") } else { 'Error' }
             $line = ($originZipStr.PadRight($col1ASP)) + ($destZipStr.PadRight($col2ASP)) + ($weightStr.PadRight($col3ASP)) + ($costStr.PadRight($col4ASP))
             $reportContent.Add($line)
        } catch { Write-Warning "Skipping result row (Origin: $($result.OriginZip)) due to formatting error: $($_.Exception.Message)"; $skippedShipmentCount++ }
    }

    # --- Report Summary ---
    $reportContent.Add("--------------------------------------------------------------------------")
    $reportContent.Add("Summary:")
    $reportContent.Add("Processed Shipments (API OK): $processedShipmentCount")
    $reportContent.Add("Skipped Shipments (API Errors/Formatting): $skippedShipmentCount")
    $avgCost = 0.0; $requiredAvgMarginPercent = $null; $avgProfitPerShipment = 0.0
    if ($processedShipmentCount -gt 0) {
        $avgCost = $totalCostValue / $processedShipmentCount
        $avgProfitPerShipment = $DesiredASP - $avgCost
        if ($DesiredASP -ne 0) { try { $requiredAvgMarginPercent = (($DesiredASP - $avgCost) / $DesiredASP) * 100.0 } catch {} }
        $reportContent.Add("Average Cost (from '$costAccountName'): $($avgCost.ToString("C2"))")
        $reportContent.Add("Desired Average Sell Price (ASP): $($DesiredASP.ToString("C2"))")
        if ($requiredAvgMarginPercent -ne $null) { $reportContent.Add("Required Avg Margin % on Cost to achieve ASP: $($requiredAvgMarginPercent.ToString("N2"))%") }
        else { $reportContent.Add("Required Avg Margin % on Cost: N/A (Likely due to $0 Desired ASP or $0 Avg Cost)") }
        $reportContent.Add("Calculated Avg Profit/Shipment at Desired ASP: $($avgProfitPerShipment.ToString("C2"))")
    } else { $reportContent.Add("No shipments processed successfully, cannot calculate averages.") }
    $reportContent.Add("--------------------------------------------------------------------------")
    $reportContent.Add("End of Report")

    # --- Save Report ---
    try { $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop; Write-Host "`nCentral Transport Required Margin for ASP Report saved successfully to: $reportFilePath" -ForegroundColor Green; return $reportFilePath }
    catch { Write-Error "Failed to save Central report file '$reportFilePath': $($_.Exception.Message)"; return $null }
}


function List-CentralPermittedKeysGUI {
    # Lists Central Transport keys permitted for a given customer profile.
    param(
        [Parameter(Mandatory=$true)][hashtable]$CustomerProfile, # Customer profile from TMS_Auth
        [Parameter(Mandatory=$true)][hashtable]$AllCentralKeys    # All loaded Central keys from TMS_Main/GUI
     )
     Write-Host "`n--- Central Transport Keys Permitted for Customer: $($CustomerProfile.CustomerName) ---" -ForegroundColor Cyan
     $allowedKeyNames = @()
     # Ensure AllowedCentralKeys exists in the customer profile (should be added by TMS_Auth)
     if ($CustomerProfile.ContainsKey('AllowedCentralKeys') -and $CustomerProfile['AllowedCentralKeys'] -is [array]) {
         $allowedKeyNames = $CustomerProfile['AllowedCentralKeys']
     }

     if ($allowedKeyNames.Count -eq 0) { Write-Host "No Central Transport keys/tariffs permitted for this customer."; return }

     # Get-PermittedKeys is a helper from TMS_Helpers.ps1
     $permittedKeys = Get-PermittedKeys -AllKeys $AllCentralKeys -AllowedKeyNames $allowedKeyNames

     if ($permittedKeys.Count -eq 0) { Write-Warning "Could not retrieve details for any permitted Central keys (check if key names in customer profile match actual key file names)."; return }

     $permittedKeys.GetEnumerator() | Sort-Object Value.Name | ForEach-Object { # Sort by the 'Name' property within the hashtable value
         $keyDetails = $_.Value
         $keyDisplayName = if ($keyDetails.ContainsKey('Name')) { $keyDetails.Name } else { $_.Key } # Fallback to filename if Name prop missing
         Write-Host "`nKey/Tariff Name: $keyDisplayName" -ForegroundColor Yellow
         if ($keyDetails -is [hashtable]) {
             $keyDetails.GetEnumerator() | Where-Object { $_.Name -ne 'Name' -and $_.Name -ne 'TariffFileName' } | Sort-Object Name | ForEach-Object {
                 Write-Host "  $($_.Name): $($_.Value)"
             }
         } else { Write-Warning "  Data for key '$keyDisplayName' is not in the expected hashtable format." }
     }
}

Write-Verbose "TMS Central Transport Functions loaded."
