# TMS_Carrier_Averitt.ps1
# Description: Contains functions specific to Averitt Express operations,
#              refactored to accept parameters for GUI use.
#              Requires TMS_Helpers.ps1 and TMS_Config.ps1 to be loaded first (by main entry script).
#              This file should be dot-sourced by the main entry script (TMS_GUI.ps1).

# Assumes helper functions like Invoke-AverittApi, Write-LoadingBar,
# Load-And-Normalize-AverittData, Get-ReportPath, Select-CsvFile,
# Select-SingleKeyEntry, Get-PermittedKeys are available from TMS_Helpers.ps1 or main script.
# Assumes config variables like $script:averittApiUri are available from TMS_Config.ps1 (via main script).

function Run-AverittComparisonReportGUI {
    param(
        [Parameter(Mandatory=$true)][hashtable]$Key1Data,
        [Parameter(Mandatory=$true)][hashtable]$Key2Data,
        [Parameter(Mandatory=$true)][string]$CsvFilePath, # Expects a detailed CSV like shipments.csv
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$UserReportsFolder
    )
    Write-Host "`nRunning Averitt Comparison Report (GUI Mode)..." -ForegroundColor Cyan
    $key1DisplayName = if ($Key1Data.ContainsKey('Name')) { $Key1Data.Name } else { $Key1Data.TariffFileName | Split-Path -Leaf }
    $key2DisplayName = if ($Key2Data.ContainsKey('Name')) { $Key2Data.Name } else { $Key2Data.TariffFileName | Split-Path -Leaf }
    Write-Host "Comparing Account: '$key1DisplayName' vs Account: '$key2DisplayName'"

    # --- Data Loading (Averitt specific) ---
    $shipments = Load-And-Normalize-AverittData -CsvPath $CsvFilePath # Expects detailed CSV
    if ($shipments -eq $null -or $shipments.Count -eq 0) {
        Write-Warning "No processable Averitt shipment data found in '$CsvFilePath'."
        return $null
    }

    # --- Report Preparation ---
    $reportContent = [System.Collections.Generic.List[string]]::new()
    $resultsData = [System.Collections.Generic.List[object]]::new()
    $key1NameSafe = $key1DisplayName -replace '[^a-zA-Z0-9_-]', ''
    $key2NameSafe = $key2DisplayName -replace '[^a-zA-Z0-9_-]', ''
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $Username -Carrier 'Averitt' -ReportType 'Comparison' -FilePrefix ($key1NameSafe + "_vs_" + $key2NameSafe)
    if (-not $reportFilePath) { return $null }

    $skippedShipmentCount = 0; $totalDifference = 0.0; $processedShipmentCount = 0
    $reportContent.Add("Averitt Comparison Report"); $reportContent.Add("User: $Username"); $reportContent.Add("Date: $(Get-Date)"); $reportContent.Add("Data File: $CsvFilePath")
    $reportContent.Add("Comparing Account: '$key1DisplayName' vs Account: '$key2DisplayName'")
    $reportContent.Add("-------------------------------------------------------------------------------------")
    $col1WidthComp = 15; $col2WidthComp = 15; $col3WidthComp = 12; $col4WidthComp = 18; $col5WidthComp = 18; $col6WidthComp = 15
    $headerLineComp = ("Origin Zip".PadRight($col1WidthComp)) + ("Dest Zip".PadRight($col2WidthComp)) + ("Total Weight".PadRight($col3WidthComp)) + ("$key1DisplayName Cost".PadRight($col4WidthComp)) + ("$key2DisplayName Cost".PadRight($col5WidthComp)) + ("Difference".PadRight($col6WidthComp))
    $reportContent.Add($headerLineComp); $reportContent.Add("-------------------------------------------------------------------------------------")

    # --- Process Shipments ---
    $totalShipmentsToProcess = $shipments.Count
    Write-Host "Processing $totalShipmentsToProcess shipments for Averitt Comparison..."
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt $totalShipmentsToProcess; $i++) {
        $shipmentDataRow = $shipments[$i]; $shipmentNumberForDisplay = $i + 1
        if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($shipmentNumberForDisplay * 100 / $totalShipmentsToProcess)) -Message "Processing Shipment $shipmentNumberForDisplay (Averitt Comp)" }

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'
        $cost1 = Invoke-AverittApi -KeyData $Key1Data -ShipmentData $shipmentDataRow
        $cost2 = Invoke-AverittApi -KeyData $Key2Data -ShipmentData $shipmentDataRow
        $VerbosePreference = $CurrentVerbosePreference

        if ($cost1 -eq $null -or $cost2 -eq $null) {
            $skippedShipmentCount++; Write-Warning "Skipping Averitt shipment $shipmentNumberForDisplay (Origin: $($shipmentDataRow.OriginPostalCode)) due to API error or no rate."
            continue
        }

        $difference = $cost1 - $cost2; $totalDifference += $difference; $processedShipmentCount++
        
        $currentWeight = 0.0
        try { # Corrected try-catch block for weight calculation
            for ($commIdx = 1; $commIdx -le 5; $commIdx++) {
                $weightKey = "Commodity${commIdx}_Weight"
                if ($shipmentDataRow.PSObject.Properties.Match($weightKey) -and (-not [string]::IsNullOrWhiteSpace($shipmentDataRow.$weightKey))) {
                    $currentWeight += [decimal]$shipmentDataRow.$weightKey
                }
            }
            # If all commodity weights are zero/missing, try a total weight column if it exists from the CSV
            if ($currentWeight -eq 0.0 -and $shipmentDataRow.PSObject.Properties.Match('Total_Weight_From_CSV')) { # Assuming a column named this might exist
                 $currentWeight = [decimal]$shipmentDataRow.Total_Weight_From_CSV
            }
        } catch { # Catch for the weight calculation try block
            Write-Warning "Could not accurately sum weights for report on shipment $shipmentNumberForDisplay (Origin: $($shipmentDataRow.OriginPostalCode)). Error: $($_.Exception.Message)"
            # $currentWeight remains 0.0 or its last valid sum
        }


        $resultsData.Add([PSCustomObject]@{
            OriginZip = $shipmentDataRow.OriginPostalCode
            DestZip = $shipmentDataRow.DestinationPostalCode
            Weight = if($currentWeight -gt 0) {$currentWeight} else {'N/A'} 
            Cost1 = $cost1
            Cost2 = $cost2
            Difference = $difference
        })
    }
    if ($useLoadingBar) { Write-Progress -Activity "Processing Averitt Comparison Shipments" -Completed }

    # --- Format Results ---
    foreach ($result in $resultsData) {
        try {
            $originZipStr = if ([string]::IsNullOrWhiteSpace($result.OriginZip)) { 'N/A' } else { $result.OriginZip }
            $destZipStr = if ([string]::IsNullOrWhiteSpace($result.DestZip)) { 'N/A' } else { $result.DestZip }
            $weightStr = if ($result.Weight -eq 'N/A' -or $null -eq $result.Weight) { 'N/A' } else { ([decimal]$result.Weight).ToString("N0") }
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
    $reportContent.Add("-------------------------------------------------------------------------------------")
    $reportContent.Add("Summary:")
    $reportContent.Add("Processed Shipments: $processedShipmentCount")
    $reportContent.Add("Skipped Shipments (API/Formatting Errors): $skippedShipmentCount")
    if ($processedShipmentCount -gt 0) {
        $avgDifference = if ($processedShipmentCount -ne 0) { $totalDifference / $processedShipmentCount } else { 0 }
        $reportContent.Add("Total Cost Difference ($key1DisplayName - $key2DisplayName): $($totalDifference.ToString("C2"))")
        $reportContent.Add("Average Cost Difference per Shipment: $($avgDifference.ToString("C2"))")
    } else { $reportContent.Add("No shipments could be processed successfully.") }
    $reportContent.Add("-------------------------------------------------------------------------------------")
    $reportContent.Add("End of Report")

    # --- Save Report ---
    try {
        $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop
        Write-Host "`nAveritt Comparison Report saved successfully to: $reportFilePath" -ForegroundColor Green
        return $reportFilePath
    } catch {
        Write-Error "Failed to save Averitt comparison report file '$reportFilePath': $($_.Exception.Message)"
        return $null
    }
}

function Run-AverittMarginReportGUI {
    param(
        [Parameter(Mandatory=$true)][hashtable]$BaseKeyData,
        [Parameter(Mandatory=$true)][hashtable]$ComparisonKeyData,
        [Parameter(Mandatory=$true)][string]$CsvFilePath, # Expects detailed Averitt CSV
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$UserReportsFolder
    )
    Write-Host "`nRunning Averitt Average Required Margin Report (GUI Mode)..." -ForegroundColor Cyan
    $baseKeyName = if ($BaseKeyData.ContainsKey('Name')) { $BaseKeyData.Name } else { $BaseKeyData.TariffFileName | Split-Path -Leaf }
    $compKeyName = if ($ComparisonKeyData.ContainsKey('Name')) { $ComparisonKeyData.Name } else { $ComparisonKeyData.TariffFileName | Split-Path -Leaf }
    Write-Host "Base Cost Account: '$baseKeyName', Comparison Cost Account: '$compKeyName'"

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

    $shipments = Load-And-Normalize-AverittData -CsvPath $CsvFilePath
    if ($shipments -eq $null -or $shipments.Count -eq 0) { Write-Warning "No processable Averitt shipment data found in '$CsvFilePath'."; return $null }

    $reportContent = [System.Collections.Generic.List[string]]::new(); $resultsData = [System.Collections.Generic.List[object]]::new()
    $safeBaseKeyName = $baseKeyName -replace '[^a-zA-Z0-9_-]', ''; $safeCompKeyName = $compKeyName -replace '[^a-zA-Z0-9_-]', ''
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $Username -Carrier 'Averitt' -ReportType 'AvgMarginCalc' -FilePrefix ($safeBaseKeyName + "_vs_" + $safeCompKeyName)
    if (-not $reportFilePath) { return $null }

    $skippedShipmentCount = 0; $processedShipmentCount = 0; $totalTargetSellPrice = 0.0; $totalCompCost = 0.0
    $reportContent.Add("Averitt Average Required Margin Calculation Report"); $reportContent.Add("User: $Username"); $reportContent.Add("Date: $(Get-Date)"); $reportContent.Add("Data File: $CsvFilePath")
    $reportContent.Add("Base Cost Account: '$baseKeyName'"); $reportContent.Add("Comparison Cost Account: '$compKeyName'"); $reportContent.Add("Margin Applied to Base Cost (from '$baseKeyName'): $($customerCurrentMarginPercent)%")
    $reportContent.Add("--------------------------------------------------------------------------------------------------------------------")
    $col1Width = 12; $col2Width = 12; $col3Width = 12; $col4Width = 15; $col5Width = 15; $col6Width = 18; $col7Width = 15; $col8Width = 15
    $headerLine = ("Origin Zip".PadRight($col1Width)) + ("Dest Zip".PadRight($col2Width)) + ("Weight".PadRight($col3Width)) + ("Base Cost".PadRight($col4Width)) + ("Comp Cost".PadRight($col5Width)) + ("Target Sell".PadRight($col6Width)) + ("Base Profit".PadRight($col7Width)) + ("Comp Profit".PadRight($col8Width))
    $reportContent.Add($headerLine); $reportContent.Add("--------------------------------------------------------------------------------------------------------------------")

    $totalShipmentsToProcess = $shipments.Count; Write-Host "Processing $totalShipmentsToProcess shipments for Averitt Avg Margin Calc..."
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt $totalShipmentsToProcess; $i++) {
        $shipmentDataRow = $shipments[$i]; $shipmentNumberForDisplay = $i + 1
        if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($shipmentNumberForDisplay * 100 / $totalShipmentsToProcess)) -Message "Processing Shipment $shipmentNumberForDisplay (Averitt Avg Margin)" }

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'
        $baseCost = Invoke-AverittApi -KeyData $BaseKeyData -ShipmentData $shipmentDataRow
        $compCost = Invoke-AverittApi -KeyData $ComparisonKeyData -ShipmentData $shipmentDataRow
        $VerbosePreference = $CurrentVerbosePreference

        if ($baseCost -eq $null -or $compCost -eq $null) { $skippedShipmentCount++; Write-Warning "Skipping Averitt shipment $shipmentNumberForDisplay (API error)."; continue }
        if ($baseCost -le 0) { $skippedShipmentCount++; Write-Warning "Skipping Averitt shipment $shipmentNumberForDisplay (zero/negative Base Cost)."; continue }

        $targetSellPrice = $null; $baseProfit = $null; $compProfit = $null
        try {
            if ((1.0 - $customerCurrentMarginDecimal) -eq 0) { throw "Division by zero (100% margin)" }
            $targetSellPrice = $baseCost / (1.0 - $customerCurrentMarginDecimal)
            $baseProfit = $targetSellPrice - $baseCost
            $compProfit = $targetSellPrice - $compCost
        } catch { Write-Warning "Skipping Averitt shipment $shipmentNumberForDisplay (calculation error: $($_.Exception.Message))"; $skippedShipmentCount++; continue }

        $processedShipmentCount++; $totalTargetSellPrice += $targetSellPrice; $totalCompCost += $compCost
        
        $currentWeight = 0.0
        try { # Corrected try-catch block for weight calculation
            for ($commIdx = 1; $commIdx -le 5; $commIdx++) { 
                $weightKey = "Commodity${commIdx}_Weight"
                if ($shipmentDataRow.PSObject.Properties.Match($weightKey) -and -not [string]::IsNullOrWhiteSpace($shipmentDataRow.$weightKey)) { 
                    $currentWeight += [decimal]$shipmentDataRow.$weightKey
                }
            }
            if ($currentWeight -eq 0.0 -and $shipmentDataRow.PSObject.Properties.Match('Total_Weight_From_CSV')) {
                 $currentWeight = [decimal]$shipmentDataRow.Total_Weight_From_CSV
            }
        } catch { # Catch for the weight calculation try block
             Write-Warning "Could not accurately sum weights for report on shipment $shipmentNumberForDisplay (Origin: $($shipmentDataRow.OriginPostalCode)). Error: $($_.Exception.Message)"
        }


        $resultsData.Add([PSCustomObject]@{
            OriginZip = $shipmentDataRow.OriginPostalCode; DestZip = $shipmentDataRow.DestinationPostalCode; Weight = if($currentWeight -gt 0) {$currentWeight} else {'N/A'}
            BaseCost = $baseCost; CompCost = $compCost; TargetSellPrice = $targetSellPrice; BaseProfit = $baseProfit; CompProfit = $compProfit
        })
    }
     if ($useLoadingBar) { Write-Progress -Activity "Processing Averitt Avg Margin Shipments" -Completed }

    foreach ($result in $resultsData) {
        try {
             $originZipStr = if ([string]::IsNullOrWhiteSpace($result.OriginZip)) { 'N/A' } else { $result.OriginZip }
             $destZipStr = if ([string]::IsNullOrWhiteSpace($result.DestZip)) { 'N/A' } else { $result.DestZip }
             $weightStr = if ($result.Weight -eq 'N/A' -or $null -eq $result.Weight) { 'N/A' } else { ([decimal]$result.Weight).ToString("N0") }
             $baseCostStr = if ($result.BaseCost -ne $null) { $result.BaseCost.ToString("C2") } else { 'Error' }
             $compCostStr = if ($result.CompCost -ne $null) { $result.CompCost.ToString("C2") } else { 'Error' }
             $targetSellPriceStr = if ($result.TargetSellPrice -ne $null) { $result.TargetSellPrice.ToString("C2") } else { 'Error' }
             $baseProfitStr = if ($result.BaseProfit -ne $null) { $result.BaseProfit.ToString("C2") } else { 'Error' }
             $compProfitStr = if ($result.CompProfit -ne $null) { $result.CompProfit.ToString("C2") } else { 'Error' }
             $line = ($originZipStr.PadRight($col1Width)) + ($destZipStr.PadRight($col2Width)) + ($weightStr.PadRight($col3Width)) + ($baseCostStr.PadRight($col4Width)) + ($compCostStr.PadRight($col5Width)) + ($targetSellPriceStr.PadRight($col6Width)) + ($baseProfitStr.PadRight($col7Width)) + ($compProfitStr.PadRight($col8Width))
             $reportContent.Add($line)
        } catch { Write-Warning "Skipping result row (Origin: $($result.OriginZip)) due to formatting error: $($_.Exception.Message)"; $skippedShipmentCount++ }
    }

    $reportContent.Add("--------------------------------------------------------------------------------------------------------------------"); $reportContent.Add("Summary:"); $reportContent.Add("Processed Shipments (API & Calc OK): $processedShipmentCount"); $reportContent.Add("Skipped Shipments (API/Calc/Formatting Errors): $skippedShipmentCount")
    if ($processedShipmentCount -gt 0) {
        $avgTargetSellPrice = $totalTargetSellPrice / $processedShipmentCount; $avgCompCost = $totalCompCost / $processedShipmentCount
        $reportContent.Add("Average Target Sell Price (from '$baseKeyName' + $customerCurrentMarginPercent%): $($avgTargetSellPrice.ToString("C2"))")
        $reportContent.Add("Average Comparison Cost (from '$compKeyName'): $($avgCompCost.ToString("C2"))")
        $avgRequiredMarginPercent = $null
        if ($avgTargetSellPrice -ne 0) { try { $avgRequiredMarginDecimal = ($avgTargetSellPrice - $avgCompCost) / $avgTargetSellPrice; $avgRequiredMarginPercent = $avgRequiredMarginDecimal * 100.0; $reportContent.Add("Average Margin Required on '$compKeyName' Cost to Match Average Target Sell Price: $($avgRequiredMarginPercent.ToString("N2"))%") } catch { $reportContent.Add("Could not calculate avg required margin.")} }
        else { $reportContent.Add("Average Target Sell Price is zero, cannot calculate average required margin.") }
    } else { $reportContent.Add("No shipments processed successfully, cannot calculate average margin.") }
    $reportContent.Add("--------------------------------------------------------------------------------------------------------------------"); $reportContent.Add("End of Report")

    try { $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop; Write-Host "`nAveritt Average Margin Calculation Report saved successfully to: $reportFilePath" -ForegroundColor Green; return $reportFilePath }
    catch { Write-Error "Failed to save Averitt report file '$reportFilePath': $($_.Exception.Message)"; return $null }
}

function Calculate-AverittMarginForASPReportGUI {
    param(
        [Parameter(Mandatory=$true)][hashtable]$CostAccountInfo,
        [Parameter(Mandatory=$true)][decimal]$DesiredASP,
        [Parameter(Mandatory=$true)][string]$CsvFilePath,
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$UserReportsFolder
    )
    Write-Host "`nRunning Averitt Required Margin for Desired ASP Report (GUI Mode)..." -ForegroundColor Cyan
    if ($DesiredASP -le 0) { Write-Error "Desired ASP must be positive."; return $null }
    $costAccountName = if ($CostAccountInfo.ContainsKey('Name')) { $CostAccountInfo.Name } else { $CostAccountInfo.TariffFileName | Split-Path -Leaf }
    Write-Host "Cost Basis Account: '$costAccountName', Desired ASP: $($DesiredASP.ToString('C2'))"

    $shipments = Load-And-Normalize-AverittData -CsvPath $CsvFilePath
    if ($shipments -eq $null -or $shipments.Count -eq 0) { Write-Warning "No processable Averitt shipment data found in '$CsvFilePath'."; return $null }

    $reportContent = [System.Collections.Generic.List[string]]::new(); $resultsData = [System.Collections.Generic.List[object]]::new()
    $safeCostKeyName = $costAccountName -replace '[^a-zA-Z0-9_-]', ''
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $Username -Carrier 'Averitt' -ReportType 'MarginForASP' -FilePrefix $safeCostKeyName
    if (-not $reportFilePath) { return $null }

    $skippedShipmentCount = 0; $processedShipmentCount = 0; $totalCostValue = 0.0
    $reportContent.Add("Averitt Required Margin for Desired ASP Report"); $reportContent.Add("User: $Username"); $reportContent.Add("Date: $(Get-Date)"); $reportContent.Add("Data File: $CsvFilePath")
    $reportContent.Add("Cost Basis Account: '$costAccountName'"); $reportContent.Add("Desired Average Sell Price (ASP): $($DesiredASP.ToString("C2"))")
    $reportContent.Add("--------------------------------------------------------------------------")
    $col1ASP = 12; $col2ASP = 12; $col3ASP = 15; $col4ASP = 15
    $headerLineASP = ("Origin Zip".PadRight($col1ASP)) + ("Dest Zip".PadRight($col2ASP)) + ("Weight".PadRight($col3ASP)) + ("Retrieved Cost".PadRight($col4ASP))
    $reportContent.Add($headerLineASP); $reportContent.Add("--------------------------------------------------------------------------")

    $totalShipmentsToProcess = $shipments.Count; Write-Host "Processing $totalShipmentsToProcess shipments for Averitt Margin for ASP..."
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt $totalShipmentsToProcess; $i++) {
        $shipmentDataRow = $shipments[$i]; $shipmentNumberForDisplay = $i + 1
        if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($shipmentNumberForDisplay * 100 / $totalShipmentsToProcess)) -Message "Processing Shipment $shipmentNumberForDisplay (Averitt Margin/ASP)" }

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'
        $costValue = Invoke-AverittApi -KeyData $CostAccountInfo -ShipmentData $shipmentDataRow
        $VerbosePreference = $CurrentVerbosePreference

        if ($costValue -eq $null) { $skippedShipmentCount++; Write-Warning "Skipping Averitt shipment $shipmentNumberForDisplay (API error)."; continue }

        $processedShipmentCount++; $totalCostValue += $costValue
        
        $currentWeight = 0.0
        try { # Corrected try-catch block for weight calculation
            for ($commIdx = 1; $commIdx -le 5; $commIdx++) { 
                $weightKey = "Commodity${commIdx}_Weight"
                if ($shipmentDataRow.PSObject.Properties.Match($weightKey) -and -not [string]::IsNullOrWhiteSpace($shipmentDataRow.$weightKey)) { 
                    $currentWeight += [decimal]$shipmentDataRow.$weightKey
                }
            }
            if ($currentWeight -eq 0.0 -and $shipmentDataRow.PSObject.Properties.Match('Total_Weight_From_CSV')) {
                 $currentWeight = [decimal]$shipmentDataRow.Total_Weight_From_CSV
            }
        } catch { # Catch for the weight calculation try block
             Write-Warning "Could not accurately sum weights for report on shipment $shipmentNumberForDisplay (Origin: $($shipmentDataRow.OriginPostalCode)). Error: $($_.Exception.Message)"
        }

        $resultsData.Add([PSCustomObject]@{
            OriginZip = $shipmentDataRow.OriginPostalCode; DestZip = $shipmentDataRow.DestinationPostalCode; Weight = if($currentWeight -gt 0) {$currentWeight} else {'N/A'}; Cost = $costValue
        })
    }
    if ($useLoadingBar) { Write-Progress -Activity "Processing Averitt Margin/ASP Shipments" -Completed }

    foreach ($result in $resultsData) {
        try {
             $originZipStr = if ([string]::IsNullOrWhiteSpace($result.OriginZip)) { 'N/A' } else { $result.OriginZip }
             $destZipStr = if ([string]::IsNullOrWhiteSpace($result.DestZip)) { 'N/A' } else { $result.DestZip }
             $weightStr = if ($result.Weight -eq 'N/A' -or $null -eq $result.Weight) { 'N/A' } else { ([decimal]$result.Weight).ToString("N0") }
             $costStr = if ($result.Cost -ne $null) { $result.Cost.ToString("C2") } else { 'Error' }
             $line = ($originZipStr.PadRight($col1ASP)) + ($destZipStr.PadRight($col2ASP)) + ($weightStr.PadRight($col3ASP)) + ($costStr.PadRight($col4ASP))
             $reportContent.Add($line)
        } catch { Write-Warning "Skipping result row (Origin: $($result.OriginZip)) due to formatting error: $($_.Exception.Message)"; $skippedShipmentCount++ }
    }

    $reportContent.Add("--------------------------------------------------------------------------"); $reportContent.Add("Summary:"); $reportContent.Add("Processed Shipments (API OK): $processedShipmentCount"); $reportContent.Add("Skipped Shipments (API Errors/Formatting): $skippedShipmentCount")
    $avgCost = 0.0; $requiredAvgMarginPercent = $null; $avgProfitPerShipment = 0.0
    if ($processedShipmentCount -gt 0) {
        $avgCost = $totalCostValue / $processedShipmentCount; $avgProfitPerShipment = $DesiredASP - $avgCost
        if ($DesiredASP -ne 0) { try { $requiredAvgMarginPercent = (($DesiredASP - $avgCost) / $DesiredASP) * 100.0 } catch {} }
        $reportContent.Add("Average Cost (from '$costAccountName'): $($avgCost.ToString("C2"))"); $reportContent.Add("Desired Average Sell Price (ASP): $($DesiredASP.ToString("C2"))")
        if ($requiredAvgMarginPercent -ne $null) { $reportContent.Add("Required Avg Margin % on Cost to achieve ASP: $($requiredAvgMarginPercent.ToString("N2"))%") }
        else { $reportContent.Add("Required Avg Margin % on Cost: N/A (Likely due to $0 Desired ASP or $0 Avg Cost)") }
        $reportContent.Add("Calculated Avg Profit/Shipment at Desired ASP: $($avgProfitPerShipment.ToString("C2"))")
    } else { $reportContent.Add("No shipments processed successfully, cannot calculate averages.") }
    $reportContent.Add("--------------------------------------------------------------------------"); $reportContent.Add("End of Report")

    try { $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop; Write-Host "`nAveritt Required Margin for ASP Report saved successfully to: $reportFilePath" -ForegroundColor Green; return $reportFilePath }
    catch { Write-Error "Failed to save Averitt report file '$reportFilePath': $($_.Exception.Message)"; return $null }
}

function List-AverittPermittedKeysGUI {
    param(
        [Parameter(Mandatory=$true)][hashtable]$CustomerProfile,
        [Parameter(Mandatory=$true)][hashtable]$AllAverittKeys
     )
     Write-Host "`n--- Averitt Keys Permitted for Customer: $($CustomerProfile.CustomerName) ---" -ForegroundColor Cyan
     $allowedKeyNames = @()
     if ($CustomerProfile.ContainsKey('AllowedAverittKeys') -and $CustomerProfile['AllowedAverittKeys'] -is [array]) {
         $allowedKeyNames = $CustomerProfile['AllowedAverittKeys']
     }

     if ($allowedKeyNames.Count -eq 0) { Write-Host "No Averitt keys/accounts permitted for this customer."; return }

     $permittedKeys = Get-PermittedKeys -AllKeys $AllAverittKeys -AllowedKeyNames $allowedKeyNames

     if ($permittedKeys.Count -eq 0) { Write-Warning "Could not retrieve details for any permitted Averitt keys."; return }

     $permittedKeys.GetEnumerator() | Sort-Object Value.Name | ForEach-Object {
         $keyDetails = $_.Value
         $keyDisplayName = if ($keyDetails.ContainsKey('Name')) { $keyDetails.Name } else { $_.Key }
         Write-Host "`nAccount/Tariff Name: $keyDisplayName" -ForegroundColor Yellow
         if ($keyDetails -is [hashtable]) {
              $keyDetails.GetEnumerator() | Where-Object { $_.Name -ne 'Name' -and $_.Name -ne 'TariffFileName' } | Sort-Object Name | ForEach-Object {
                   $displayValue = $_.Value
                   if ($_.Name -eq 'APIKey' -and $displayValue -is [string] -and $displayValue.Length -gt 8) {
                       $displayValue = $displayValue.Substring(0, 4) + '...' + $displayValue.Substring($displayValue.Length - 4)
                   }
                   Write-Host "  $($_.Name): $displayValue"
              }
         }
         else { Write-Warning "  Data for account '$keyDisplayName' is not in the expected hashtable format." }
     }
}

Write-Verbose "TMS Averitt Carrier Functions loaded."
