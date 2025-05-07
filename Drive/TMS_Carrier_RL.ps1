# TMS_Carrier_RL.ps1
# Description: Contains functions specific to R+L Carriers operations,
#              refactored to accept parameters for GUI use.
#              Requires TMS_Helpers.ps1 and TMS_Config.ps1 to be loaded first (by main entry script).
#              This file should be dot-sourced by the main entry script (TMS_GUI.ps1).

# Assumes helper functions like Invoke-RLApi, Write-LoadingBar,
# Load-And-Normalize-RLData, Get-ReportPath, Select-CsvFile,
# Select-SingleKeyEntry, Get-PermittedKeys are available from TMS_Helpers.ps1 or main script.
# Assumes config variables like $script:rlApiUri are available from TMS_Config.ps1 (via main script).

function Run-RLComparisonReportGUI {
    # GUI VERSION: Generates a report comparing costs between two selected R+L accounts.
    param(
        [Parameter(Mandatory=$true)][hashtable]$Key1Data, # Pass the actual hashtable for Key 1
        [Parameter(Mandatory=$true)][hashtable]$Key2Data, # Pass the actual hashtable for Key 2
        [Parameter(Mandatory=$true)][string]$CsvFilePath,
        [Parameter(Mandatory=$true)][string]$Username,    # Broker Username
        [Parameter(Mandatory=$true)][string]$UserReportsFolder # Broker's Base reports folder
    )
    Write-Host "`nRunning R+L Comparison Report (GUI Mode)..." -ForegroundColor Cyan
    $key1DisplayName = if ($Key1Data.ContainsKey('Name')) { $Key1Data.Name } else { $Key1Data.TariffFileName | Split-Path -Leaf }
    $key2DisplayName = if ($Key2Data.ContainsKey('Name')) { $Key2Data.Name } else { $Key2Data.TariffFileName | Split-Path -Leaf }
    Write-Host "Comparing Account: '$key1DisplayName' vs Account: '$key2DisplayName'"

    # --- Data Loading ---
    # Load-And-Normalize-RLData expects a CSV that contains at least Origin/Dest Zip, Weight, Class, and City/State for R+L.
    $shipments = Load-And-Normalize-RLData -CsvPath $CsvFilePath
    if ($shipments -eq $null -or $shipments.Count -eq 0) {
        Write-Warning "No processable R+L shipment data found in '$CsvFilePath'."
        return $null # Indicate failure
    }

    # --- Report Preparation ---
    $reportContent = [System.Collections.Generic.List[string]]::new()
    $resultsData = [System.Collections.Generic.List[object]]::new()
    $key1NameSafe = $key1DisplayName -replace '[^a-zA-Z0-9_-]', ''
    $key2NameSafe = $key2DisplayName -replace '[^a-zA-Z0-9_-]', ''
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $Username -Carrier 'RL' -ReportType 'Comparison' -FilePrefix ($key1NameSafe + "_vs_" + $key2NameSafe)
    if (-not $reportFilePath) { return $null }

    $skippedShipmentCount = 0; $totalDifference = 0.0; $processedShipmentCount = 0
    $reportContent.Add("R+L Carriers Comparison Report"); $reportContent.Add("User: $Username"); $reportContent.Add("Date: $(Get-Date)"); $reportContent.Add("Data File: $CsvFilePath")
    $reportContent.Add("Comparing Account: '$key1DisplayName' vs Account: '$key2DisplayName'")
    $reportContent.Add("----------------------------------------------------------------------")
    $col1WidthComp = 15; $col2WidthComp = 15; $col3WidthComp = 10; $col4WidthComp = 18; $col5WidthComp = 18; $col6WidthComp = 15
    $headerLineComp = ("Origin Zip".PadRight($col1WidthComp)) + ("Dest Zip".PadRight($col2WidthComp)) + ("Weight".PadRight($col3WidthComp)) + ("$key1DisplayName Cost".PadRight($col4WidthComp)) + ("$key2DisplayName Cost".PadRight($col5WidthComp)) + ("Difference".PadRight($col6WidthComp))
    $reportContent.Add($headerLineComp); $reportContent.Add("----------------------------------------------------------------------")

    # --- Process Shipments ---
    $totalShipmentsToProcess = $shipments.Count
    Write-Host "Processing $totalShipmentsToProcess shipments for R+L Comparison..."
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt $totalShipmentsToProcess; $i++) {
        $shipment = $shipments[$i]; $shipmentNumberForDisplay = $i + 1 # $shipment is a PSCustomObject from Load-And-Normalize-RLData
        if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($shipmentNumberForDisplay * 100 / $totalShipmentsToProcess)) -Message "Processing Shipment $shipmentNumberForDisplay (R+L Comp)" }

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'
        # Invoke-RLApi requires OriginZip, DestinationZip, Weight, Class, KeyData, and optionally ShipmentDetails for more fields
        $cost1 = Invoke-RLApi -OriginZip $shipment.OriginZip -DestinationZip $shipment.DestinationZip `
                                -Weight $shipment.Weight -Class $shipment.Class `
                                -KeyData $Key1Data -ShipmentDetails $shipment # Pass the whole normalized row as ShipmentDetails

        $cost2 = Invoke-RLApi -OriginZip $shipment.OriginZip -DestinationZip $shipment.DestinationZip `
                                -Weight $shipment.Weight -Class $shipment.Class `
                                -KeyData $Key2Data -ShipmentDetails $shipment
        $VerbosePreference = $CurrentVerbosePreference

        if ($cost1 -eq $null -or $cost2 -eq $null) {
            $skippedShipmentCount++; Write-Warning "Skipping R+L shipment $shipmentNumberForDisplay (Origin: $($shipment.OriginZip)) due to API error or no rate."
            continue
        }

        $difference = $cost1 - $cost2; $totalDifference += $difference; $processedShipmentCount++
        $resultsData.Add([PSCustomObject]@{
            OriginZip = $shipment.OriginZip; DestZip = $shipment.DestinationZip; Weight = $shipment.Weight # Already decimal
            Cost1 = $cost1; Cost2 = $cost2; Difference = $difference
        })
    }
    if ($useLoadingBar) { Write-Progress -Activity "Processing R+L Comparison Shipments" -Completed }

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
    try { $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop; Write-Host "`nR+L Comparison Report saved successfully to: $reportFilePath" -ForegroundColor Green; return $reportFilePath }
    catch { Write-Error "Failed to save R+L comparison report file '$reportFilePath': $($_.Exception.Message)"; return $null }
}


function Run-RLMarginReportGUI {
    # GUI VERSION: Calculates the average margin required for a 'Comparison' R+L account cost to match the average target sell price.
     param(
        [Parameter(Mandatory=$true)][hashtable]$BaseKeyData,
        [Parameter(Mandatory=$true)][hashtable]$ComparisonKeyData,
        [Parameter(Mandatory=$true)][string]$CsvFilePath,
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$UserReportsFolder
     )
    Write-Host "`nRunning R+L Average Required Margin Report (GUI Mode)..." -ForegroundColor Cyan
    $baseKeyName = if ($BaseKeyData.ContainsKey('Name')) { $BaseKeyData.Name } else { $BaseKeyData.TariffFileName | Split-Path -Leaf }
    $compKeyName = if ($ComparisonKeyData.ContainsKey('Name')) { $ComparisonKeyData.Name } else { $ComparisonKeyData.TariffFileName | Split-Path -Leaf }
    Write-Host "Base Cost Account: '$baseKeyName', Comparison Cost Account: '$compKeyName'"

    # --- Get Margin from Base Key File ---
    $customerCurrentMarginPercent = $null
    if ($BaseKeyData.ContainsKey('MarginPercent')) {
        try { $marginValue = [double]$BaseKeyData.MarginPercent; if ($marginValue -ge 0 -and $marginValue -lt 100) { $customerCurrentMarginPercent = $marginValue } else { Write-Warning "Margin '$($BaseKeyData.MarginPercent)' in '$baseKeyName' invalid." } }
        catch { Write-Warning "Invalid MarginPercent value '$($BaseKeyData.MarginPercent)' in '$baseKeyName'." }
    } else { Write-Warning "'MarginPercent' not found in base key '$baseKeyName'." }

    if ($customerCurrentMarginPercent -eq $null) {
        if ($Global:PSBoundParameters.ContainsKey('DefaultMarginPercentage') -and $Global:DefaultMarginPercentage -ge 0 -and $Global:DefaultMarginPercentage -lt 100) {
            $customerCurrentMarginPercent = $Global:DefaultMarginPercentage; Write-Warning "Using global default margin: $customerCurrentMarginPercent%"
        } else { Write-Error "Valid margin not found for '$baseKeyName' and no valid DefaultMarginPercentage configured. Cannot proceed."; return $null }
    }
    $customerCurrentMarginDecimal = [decimal]$customerCurrentMarginPercent / 100.0
    Write-Host "Using Margin from Base Account '$baseKeyName': $customerCurrentMarginPercent%" -ForegroundColor Cyan

    # --- Data Loading ---
    $shipments = Load-And-Normalize-RLData -CsvPath $CsvFilePath
    if ($shipments -eq $null -or $shipments.Count -eq 0) { Write-Warning "No processable R+L shipment data found in '$CsvFilePath'."; return $null }

    # --- Report Preparation ---
    $reportContent = [System.Collections.Generic.List[string]]::new(); $resultsData = [System.Collections.Generic.List[object]]::new()
    $safeBaseKeyName = $baseKeyName -replace '[^a-zA-Z0-9_-]', ''; $safeCompKeyName = $compKeyName -replace '[^a-zA-Z0-9_-]', ''
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $Username -Carrier 'RL' -ReportType 'AvgMarginCalc' -FilePrefix ($safeBaseKeyName + "_vs_" + $safeCompKeyName)
    if (-not $reportFilePath) { return $null }

    $skippedShipmentCount = 0; $processedShipmentCount = 0; $totalTargetSellPrice = 0.0; $totalCompCost = 0.0
    $reportContent.Add("R+L Carriers Average Required Margin Calculation Report"); $reportContent.Add("User: $Username"); $reportContent.Add("Date: $(Get-Date)"); $reportContent.Add("Data File: $CsvFilePath")
    $reportContent.Add("Base Cost Account: '$baseKeyName'"); $reportContent.Add("Comparison Cost Account: '$compKeyName'"); $reportContent.Add("Margin Applied to Base Cost (from '$baseKeyName'): $($customerCurrentMarginPercent)%")
    $reportContent.Add("--------------------------------------------------------------------------------------------------------------------")
    $col1Width = 12; $col2Width = 12; $col3Width = 10; $col4Width = 15; $col5Width = 15; $col6Width = 18; $col7Width = 15; $col8Width = 15
    $headerLine = ("Origin Zip".PadRight($col1Width)) + ("Dest Zip".PadRight($col2Width)) + ("Weight".PadRight($col3Width)) + ("Base Cost".PadRight($col4Width)) + ("Comp Cost".PadRight($col5Width)) + ("Target Sell".PadRight($col6Width)) + ("Base Profit".PadRight($col7Width)) + ("Comp Profit".PadRight($col8Width))
    $reportContent.Add($headerLine); $reportContent.Add("--------------------------------------------------------------------------------------------------------------------")

    # --- Process Shipments ---
    $totalShipmentsToProcess = $shipments.Count; Write-Host "Processing $totalShipmentsToProcess shipments for R+L Avg Margin Calc..."
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt $totalShipmentsToProcess; $i++) {
        $shipment = $shipments[$i]; $shipmentNumberForDisplay = $i + 1
        if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($shipmentNumberForDisplay * 100 / $totalShipmentsToProcess)) -Message "Processing Shipment $shipmentNumberForDisplay (R+L Avg Margin)" }

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'
        $baseCost = Invoke-RLApi -OriginZip $shipment.OriginZip -DestinationZip $shipment.DestinationZip `
                                 -Weight $shipment.Weight -Class $shipment.Class `
                                 -KeyData $BaseKeyData -ShipmentDetails $shipment
        $compCost = Invoke-RLApi -OriginZip $shipment.OriginZip -DestinationZip $shipment.DestinationZip `
                                 -Weight $shipment.Weight -Class $shipment.Class `
                                 -KeyData $ComparisonKeyData -ShipmentDetails $shipment
        $VerbosePreference = $CurrentVerbosePreference

        if ($baseCost -eq $null -or $compCost -eq $null) { $skippedShipmentCount++; Write-Warning "Skipping R+L shipment $shipmentNumberForDisplay (API error)."; continue }
        if ($baseCost -le 0) { $skippedShipmentCount++; Write-Warning "Skipping R+L shipment $shipmentNumberForDisplay (zero/negative Base Cost)."; continue }

        $targetSellPrice = $null; $baseProfit = $null; $compProfit = $null
        try {
            if ((1.0 - $customerCurrentMarginDecimal) -eq 0) { throw "Division by zero (100% margin)" }
            $targetSellPrice = $baseCost / (1.0 - $customerCurrentMarginDecimal)
            $baseProfit = $targetSellPrice - $baseCost
            $compProfit = $targetSellPrice - $compCost
        } catch { Write-Warning "Skipping R+L shipment $shipmentNumberForDisplay (calculation error: $($_.Exception.Message))"; $skippedShipmentCount++; continue }

        $processedShipmentCount++; $totalTargetSellPrice += $targetSellPrice; $totalCompCost += $compCost
        $resultsData.Add([PSCustomObject]@{
            OriginZip = $shipment.OriginZip; DestZip = $shipment.DestinationZip; Weight = $shipment.Weight
            BaseCost = $baseCost; CompCost = $compCost; TargetSellPrice = $targetSellPrice; BaseProfit = $baseProfit; CompProfit = $compProfit
        })
    }
     if ($useLoadingBar) { Write-Progress -Activity "Processing R+L Avg Margin Shipments" -Completed }

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
    $reportContent.Add("--------------------------------------------------------------------------------------------------------------------")
    $reportContent.Add("Summary:")
    $reportContent.Add("Processed Shipments (API & Calc OK): $processedShipmentCount")
    $reportContent.Add("Skipped Shipments (API/Calc/Formatting Errors): $skippedShipmentCount")
    if ($processedShipmentCount -gt 0) {
        $avgTargetSellPrice = $totalTargetSellPrice / $processedShipmentCount; $avgCompCost = $totalCompCost / $processedShipmentCount
        $reportContent.Add("Average Target Sell Price (from '$baseKeyName' + $customerCurrentMarginPercent%): $($avgTargetSellPrice.ToString("C2"))")
        $reportContent.Add("Average Comparison Cost (from '$compKeyName'): $($avgCompCost.ToString("C2"))")
        $avgRequiredMarginPercent = $null
        if ($avgTargetSellPrice -ne 0) { try { $avgRequiredMarginDecimal = ($avgTargetSellPrice - $avgCompCost) / $avgTargetSellPrice; $avgRequiredMarginPercent = $avgRequiredMarginDecimal * 100.0; $reportContent.Add("Average Margin Required on '$compKeyName' Cost to Match Average Target Sell Price: $($avgRequiredMarginPercent.ToString("N2"))%") } catch { $reportContent.Add("Could not calculate avg required margin.")} }
        else { $reportContent.Add("Average Target Sell Price is zero, cannot calculate average required margin.") }
    } else { $reportContent.Add("No shipments processed successfully, cannot calculate average margin.") }
    $reportContent.Add("--------------------------------------------------------------------------------------------------------------------")
    $reportContent.Add("End of Report")

    # --- Save Report ---
    try { $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop; Write-Host "`nR+L Average Margin Calculation Report saved successfully to: $reportFilePath" -ForegroundColor Green; return $reportFilePath }
    catch { Write-Error "Failed to save R+L report file '$reportFilePath': $($_.Exception.Message)"; return $null }
}


function Calculate-RLMarginForASPReportGUI {
    # GUI VERSION: Calculates the required margin for a specific R+L account to meet a desired ASP.
    param(
        [Parameter(Mandatory=$true)][hashtable]$CostAccountInfo, # Pass the actual hashtable for the selected key
        [Parameter(Mandatory=$true)][decimal]$DesiredASP,      # Pass the desired ASP value
        [Parameter(Mandatory=$true)][string]$CsvFilePath,
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$UserReportsFolder
    )
    Write-Host "`nRunning R+L Required Margin for Desired ASP Report (GUI Mode)..." -ForegroundColor Cyan
    if ($DesiredASP -le 0) { Write-Error "Desired ASP must be positive."; return $null }
    $costAccountName = if ($CostAccountInfo.ContainsKey('Name')) { $CostAccountInfo.Name } else { $CostAccountInfo.TariffFileName | Split-Path -Leaf }
    Write-Host "Cost Basis Account: '$costAccountName', Desired ASP: $($DesiredASP.ToString('C2'))"

    # --- Data Loading ---
    $shipments = Load-And-Normalize-RLData -CsvPath $CsvFilePath
    if ($shipments -eq $null -or $shipments.Count -eq 0) { Write-Warning "No processable R+L shipment data found in '$CsvFilePath'."; return $null }

    # --- Report Preparation ---
    $reportContent = [System.Collections.Generic.List[string]]::new(); $resultsData = [System.Collections.Generic.List[object]]::new()
    $safeCostKeyName = $costAccountName -replace '[^a-zA-Z0-9_-]', ''
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $Username -Carrier 'RL' -ReportType 'MarginForASP' -FilePrefix $safeCostKeyName
    if (-not $reportFilePath) { return $null }

    $skippedShipmentCount = 0; $processedShipmentCount = 0; $totalCostValue = 0.0
    $reportContent.Add("R+L Carriers Required Margin for Desired ASP Report"); $reportContent.Add("User: $Username"); $reportContent.Add("Date: $(Get-Date)"); $reportContent.Add("Data File: $CsvFilePath")
    $reportContent.Add("Cost Basis Account: '$costAccountName'"); $reportContent.Add("Desired Average Sell Price (ASP): $($DesiredASP.ToString("C2"))")
    $reportContent.Add("--------------------------------------------------------------------------")
    $col1ASP = 12; $col2ASP = 12; $col3ASP = 15; $col4ASP = 15
    $headerLineASP = ("Origin Zip".PadRight($col1ASP)) + ("Dest Zip".PadRight($col2ASP)) + ("Weight".PadRight($col3ASP)) + ("Retrieved Cost".PadRight($col4ASP))
    $reportContent.Add($headerLineASP); $reportContent.Add("--------------------------------------------------------------------------")

    # --- Process Shipments ---
    $totalShipmentsToProcess = $shipments.Count; Write-Host "Processing $totalShipmentsToProcess shipments for R+L Margin for ASP..."
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt $totalShipmentsToProcess; $i++) {
        $shipment = $shipments[$i]; $shipmentNumberForDisplay = $i + 1
        if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($shipmentNumberForDisplay * 100 / $totalShipmentsToProcess)) -Message "Processing Shipment $shipmentNumberForDisplay (R+L Margin/ASP)" }

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'
        $costValue = Invoke-RLApi -OriginZip $shipment.OriginZip -DestinationZip $shipment.DestinationZip `
                                  -Weight $shipment.Weight -Class $shipment.Class `
                                  -KeyData $CostAccountInfo -ShipmentDetails $shipment
        $VerbosePreference = $CurrentVerbosePreference

        if ($costValue -eq $null) { $skippedShipmentCount++; Write-Warning "Skipping R+L shipment $shipmentNumberForDisplay (API error)."; continue }

        $processedShipmentCount++; $totalCostValue += $costValue
        $resultsData.Add([PSCustomObject]@{
            OriginZip = $shipment.OriginZip; DestZip = $shipment.DestinationZip; Weight = $shipment.Weight; Cost = $costValue
        })
    }
    if ($useLoadingBar) { Write-Progress -Activity "Processing R+L Margin/ASP Shipments" -Completed }

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
        $avgCost = $totalCostValue / $processedShipmentCount; $avgProfitPerShipment = $DesiredASP - $avgCost
        if ($DesiredASP -ne 0) { try { $requiredAvgMarginPercent = (($DesiredASP - $avgCost) / $DesiredASP) * 100.0 } catch {} }
        $reportContent.Add("Average Cost (from '$costAccountName'): $($avgCost.ToString("C2"))"); $reportContent.Add("Desired Average Sell Price (ASP): $($DesiredASP.ToString("C2"))")
        if ($requiredAvgMarginPercent -ne $null) { $reportContent.Add("Required Avg Margin % on Cost to achieve ASP: $($requiredAvgMarginPercent.ToString("N2"))%") }
        else { $reportContent.Add("Required Avg Margin % on Cost: N/A (Likely due to $0 Desired ASP or $0 Avg Cost)") }
        $reportContent.Add("Calculated Avg Profit/Shipment at Desired ASP: $($avgProfitPerShipment.ToString("C2"))")
    } else { $reportContent.Add("No shipments processed successfully, cannot calculate averages.") }
    $reportContent.Add("--------------------------------------------------------------------------")
    $reportContent.Add("End of Report")

    # --- Save Report ---
    try { $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop; Write-Host "`nR+L Required Margin for ASP Report saved successfully to: $reportFilePath" -ForegroundColor Green; return $reportFilePath }
    catch { Write-Error "Failed to save R+L report file '$reportFilePath': $($_.Exception.Message)"; return $null }
}


function List-RLPermittedKeysGUI {
    # Lists R+L keys permitted for a given customer profile.
    param(
        [Parameter(Mandatory=$true)][hashtable]$CustomerProfile,
        [Parameter(Mandatory=$true)][hashtable]$AllRLKeys
    )
    Write-Host "`n--- R+L Keys Permitted for Customer: $($CustomerProfile.CustomerName) ---" -ForegroundColor Cyan
    $allowedKeyNames = @()
    if ($CustomerProfile.ContainsKey('AllowedRLKeys') -and $CustomerProfile['AllowedRLKeys'] -is [array]) {
        $allowedKeyNames = $CustomerProfile['AllowedRLKeys']
    }

    if ($allowedKeyNames.Count -eq 0) { Write-Host "No R+L keys/accounts permitted for this customer."; return }

    $permittedKeys = Get-PermittedKeys -AllKeys $AllRLKeys -AllowedKeyNames $allowedKeyNames

    if ($permittedKeys.Count -eq 0) { Write-Warning "Could not retrieve details for any permitted R+L keys."; return }

    $permittedKeys.GetEnumerator() | Sort-Object Value.Name | ForEach-Object {
        $keyDetails = $_.Value
        $keyDisplayName = if ($keyDetails.ContainsKey('Name')) { $keyDetails.Name } else { $_.Key }
        Write-Host "`nAccount Name: $keyDisplayName" -ForegroundColor Yellow
        if ($keyDetails -is [hashtable]) {
            $keyDetails.GetEnumerator() | Where-Object { $_.Name -ne 'Name' -and $_.Name -ne 'TariffFileName' } | Sort-Object Name | ForEach-Object {
                $displayValue = $_.Value
                # Example: Mask APIKey partially if desired for display
                if ($_.Name -eq 'APIKey' -and $displayValue -is [string] -and $displayValue.Length -gt 8) {
                    $displayValue = $displayValue.Substring(0, 4) + '...' + $displayValue.Substring($displayValue.Length - 4)
                }
                Write-Host "  $($_.Name): $displayValue"
            }
        } else { Write-Warning "  Data for account '$keyDisplayName' is not in the expected hashtable format." }
    }
}

Write-Verbose "TMS R+L Carrier Functions loaded."
