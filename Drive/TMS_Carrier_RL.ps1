# TMS_Carrier_RL.ps1
# Description: Contains functions specific to R+L Carriers operations,
#              refactored to accept parameters for GUI use.
#              Requires TMS_Helpers.ps1 and TMS_Config.ps1 to be loaded first.
#              This file should be dot-sourced by the main script.

# Assumes helper functions like Invoke-RLApi, Write-LoadingBar,
# Open-FileExplorer, Load-And-Normalize-RLData, Get-ReportPath are available from TMS_Helpers.ps1
# Assumes config variables like $script:rlApiUri are available from TMS_Config.ps1

function Run-RLComparisonReportGUI {
    # GUI VERSION: Generates a report comparing costs between two selected R+L keys/tariffs.
    param(
        [Parameter(Mandatory=$true)][hashtable]$Key1Data, 
        [Parameter(Mandatory=$true)][hashtable]$Key2Data, 
        [Parameter(Mandatory=$true)][string]$CsvFilePath,
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$UserReportsFolder 
    )
    Write-Host "`nRunning R+L Comparison Report (GUI Mode)..." -ForegroundColor Cyan
    Write-Host "Comparing: '$($Key1Data.Name)' vs '$($Key2Data.Name)'"

    $shipments = Load-And-Normalize-RLData -CsvPath $CsvFilePath # Use R+L specific normalization
    if ($shipments -eq $null -or $shipments.Count -eq 0) {
        Write-Warning "No processable shipment data found in '$CsvFilePath'."
        return $null 
    }

    $reportContent = [System.Collections.Generic.List[string]]::new()
    $resultsData = [System.Collections.Generic.List[object]]::new()
    $key1NameSafe = $Key1Data.Name -replace '[^a-zA-Z0-9_-]', ''
    $key2NameSafe = $Key2Data.Name -replace '[^a-zA-Z0-9_-]', ''
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $Username -Carrier 'RL' -ReportType 'Comparison' -FilePrefix ($key1NameSafe + "_vs_" + $key2NameSafe)
    if (-not $reportFilePath) { return $null }

    $skippedShipmentCount = 0; $totalDifference = 0.0; $processedShipmentCount = 0
    $reportContent.Add("R+L Carriers Comparison Report"); $reportContent.Add("User: $Username"); $reportContent.Add("Date: $(Get-Date)"); $reportContent.Add("Data File: $CsvFilePath")
    $reportContent.Add("Comparing: '$($Key1Data.Name)' vs '$($Key2Data.Name)'")
    $reportContent.Add("-----------------------------------------------------------------------------------------------------------------") 
    $col1WidthComp = 12; $col2WidthComp = 12; $col3WidthComp = 10; $col4WidthComp = 15; $col5WidthComp = 15; $col6WidthComp = 15; $col7WidthComp = 15
    $headerLineComp = ("Origin Zip".PadRight($col1WidthComp)) + ("Dest Zip".PadRight($col2WidthComp)) + ("Weight".PadRight($col3WidthComp)) + ("Class".PadRight($col4WidthComp)) + ("$($Key1Data.Name) Cost".PadRight($col5WidthComp)) + ("$($Key2Data.Name) Cost".PadRight($col6WidthComp)) + ("Difference".PadRight($col7WidthComp))
    $reportContent.Add($headerLineComp); $reportContent.Add("-----------------------------------------------------------------------------------------------------------------")

    $totalShipments = $shipments.Count
    Write-Host "Processing $totalShipments shipments..."
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt $totalShipments; $i++) {
        $shipment = $shipments[$i]; $shipmentNumber = $i + 1
        if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($shipmentNumber * 100 / $totalShipments)) -Message "Processing Shipment $shipmentNumber" }

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'
        # Pass the full $shipment object to Invoke-RLApi as it might need more details than just the few common ones
        $cost1 = Invoke-RLApi -OriginZip $shipment.OriginZip -DestinationZip $shipment.DestinationZip -Weight $shipment.Weight -Class $shipment.Class -KeyData $Key1Data -ShipmentDetails $shipment
        $cost2 = Invoke-RLApi -OriginZip $shipment.OriginZip -DestinationZip $shipment.DestinationZip -Weight $shipment.Weight -Class $shipment.Class -KeyData $Key2Data -ShipmentDetails $shipment
        $VerbosePreference = $CurrentVerbosePreference

        if ($cost1 -eq $null -or $cost2 -eq $null) { $skippedShipmentCount++; Write-Warning "Skipping shipment $shipmentNumber (Origin: $($shipment.OriginZip)) due to API error."; continue }

        $difference = $cost1 - $cost2; $totalDifference += $difference; $processedShipmentCount++
        $resultsData.Add([PSCustomObject]@{ 
            OriginZip = $shipment.OriginZip; DestZip = $shipment.DestinationZip; 
            Weight = $shipment.Weight; Class = $shipment.Class;
            Cost1 = $cost1; Cost2 = $cost2; Difference = $difference 
        })
    }
    if ($useLoadingBar) { Write-Progress -Activity "Processing Shipments" -Completed }

    foreach ($result in $resultsData) {
        try {
            $originZipStr = if ([string]::IsNullOrWhiteSpace($result.OriginZip)) { 'N/A' } else { $result.OriginZip }; $destZipStr = if ([string]::IsNullOrWhiteSpace($result.DestZip)) { 'N/A' } else { $result.DestZip }
            $weightStr = if ($null -eq $result.Weight) { 'N/A' } else { $result.Weight.ToString("N0") }; $classStr = if ([string]::IsNullOrWhiteSpace($result.Class)) { 'N/A' } else { $result.Class }
            $cost1Str = if ($result.Cost1 -ne $null) { $result.Cost1.ToString("C2") } else { 'Error' }; $cost2Str = if ($result.Cost2 -ne $null) { $result.Cost2.ToString("C2") } else { 'Error' }
            $diffStr = if ($result.Difference -ne $null) { $result.Difference.ToString("C2") } else { 'Error' }
            $line = ($originZipStr.PadRight($col1WidthComp)) + ($destZipStr.PadRight($col2WidthComp)) + ($weightStr.PadRight($col3WidthComp)) + ($classStr.PadRight($col4WidthComp)) + ($cost1Str.PadRight($col5WidthComp)) + ($cost2Str.PadRight($col6WidthComp)) + ($diffStr.PadRight($col7WidthComp))
            $reportContent.Add($line)
        } catch { Write-Warning "Skipping result row (Origin: $($result.OriginZip)) due to formatting error: $($_.Exception.Message)"; $skippedShipmentCount++; if($processedShipmentCount -gt 0){$processedShipmentCount--}; if($result.Difference -ne $null){$totalDifference -= $result.Difference} }
    }

    $reportContent.Add("-----------------------------------------------------------------------------------------------------------------"); $reportContent.Add("Summary:"); $reportContent.Add("Processed Shipments: $processedShipmentCount"); $reportContent.Add("Skipped Shipments (API/Formatting Errors): $skippedShipmentCount")
    if ($processedShipmentCount -gt 0) { $avgDifference = if ($processedShipmentCount -ne 0) { $totalDifference / $processedShipmentCount } else { 0 }; $reportContent.Add("Total Cost Difference ($($Key1Data.Name) - $($Key2Data.Name)): $($totalDifference.ToString("C2"))"); $reportContent.Add("Average Cost Difference per Shipment: $($avgDifference.ToString("C2"))") }
    else { $reportContent.Add("No shipments could be processed successfully.") }
    $reportContent.Add("-----------------------------------------------------------------------------------------------------------------"); $reportContent.Add("End of Report")

    try { $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop; Write-Host "`nComparison Report saved successfully to: $reportFilePath" -ForegroundColor Green; return $reportFilePath } 
    catch { Write-Error "Failed to save report file '$reportFilePath': $($_.Exception.Message)"; return $null } 
}


function Run-RLMarginReportGUI {
     param(
        [Parameter(Mandatory=$true)][hashtable]$BaseKeyData,      
        [Parameter(Mandatory=$true)][hashtable]$ComparisonKeyData, 
        [Parameter(Mandatory=$true)][string]$CsvFilePath,
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$UserReportsFolder
     )
    Write-Host "`nRunning R+L Average Required Margin Report (GUI Mode)..." -ForegroundColor Cyan
    $baseKeyName = $BaseKeyData.Name; $compKeyName = $ComparisonKeyData.Name
    Write-Host "Base Cost Key: '$baseKeyName', Comparison Cost Key: '$compKeyName'"

    $customerCurrentMarginPercent = $null
    if ($BaseKeyData.ContainsKey('MarginPercent')) { try { $marginValue = [double]$BaseKeyData.MarginPercent; if ($marginValue -ge 0 -and $marginValue -lt 100) { $customerCurrentMarginPercent = $marginValue } else { Write-Warning "Margin '$($BaseKeyData.MarginPercent)' in '$baseKeyName' invalid (0-99.9)." } } catch { Write-Warning "Invalid Margin value '$($BaseKeyData.MarginPercent)' in '$baseKeyName'." } } else { Write-Warning "'MarginPercent' not found in '$baseKeyName'." }
    if ($customerCurrentMarginPercent -eq $null) { if ($Global:PSBoundParameters.ContainsKey('DefaultMarginPercentage') -and $Global:DefaultMarginPercentage -ge 0 -and $Global:DefaultMarginPercentage -lt 100) { $customerCurrentMarginPercent = $Global:DefaultMarginPercentage; Write-Warning "Using default margin: $customerCurrentMarginPercent%" } else { Write-Error "Valid margin not found for '$baseKeyName' and no valid DefaultMarginPercentage configured."; return $null } }
    $customerCurrentMarginDecimal = [decimal]$customerCurrentMarginPercent / 100.0
    Write-Host "Using Margin from Base Key '$baseKeyName': $customerCurrentMarginPercent%" -ForegroundColor Cyan

    $shipments = Load-And-Normalize-RLData -CsvPath $CsvFilePath # Use R+L specific normalization
    if ($shipments -eq $null -or $shipments.Count -eq 0) { Write-Warning "No processable shipment data found in '$CsvFilePath'."; return $null }

    $reportContent = [System.Collections.Generic.List[string]]::new(); $resultsData = [System.Collections.Generic.List[object]]::new()
    $safeBaseKeyName = $baseKeyName -replace '[^a-zA-Z0-9_-]', ''; $safeCompKeyName = $compKeyName -replace '[^a-zA-Z0-9_-]', ''
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $Username -Carrier 'RL' -ReportType 'AvgMarginCalc' -FilePrefix ($safeBaseKeyName + "_vs_" + $safeCompKeyName)
    if (-not $reportFilePath) { return $null }

    $skippedShipmentCount = 0; $processedShipmentCount = 0; $totalTargetSellPrice = 0.0; $totalCompCost = 0.0
    $reportContent.Add("R+L Carriers Average Required Margin Calculation Report"); $reportContent.Add("User: $Username"); $reportContent.Add("Date: $(Get-Date)"); $reportContent.Add("Data File: $CsvFilePath")
    $reportContent.Add("Base Cost Key: '$baseKeyName'"); $reportContent.Add("Comparison Cost Key: '$compKeyName'"); $reportContent.Add("Margin Applied to Base Cost (from '$baseKeyName'): $($customerCurrentMarginPercent)%")
    $reportContent.Add("------------------------------------------------------------------------------------------------------------------------------------") 
    $col1Width = 10; $col2Width = 10; $col3Width = 12; $col4Width = 12; $col5Width = 15; $col6Width = 15; $col7Width = 15; $col8Width = 15; $col9Width = 15
    $headerLine = ("Origin Zip".PadRight($col1Width)) + ("Dest Zip".PadRight($col2Width)) + ("Weight".PadRight($col3Width)) + ("Class".PadRight($col4Width)) + ("Base Cost".PadRight($col5Width)) + ("Comp Cost".PadRight($col6Width)) + ("Target Sell".PadRight($col7Width)) + ("Base Profit".PadRight($col8Width)) + ("Comp Profit".PadRight($col9Width))
    $reportContent.Add($headerLine); $reportContent.Add("------------------------------------------------------------------------------------------------------------------------------------") 

    $totalShipments = $shipments.Count; Write-Host "Processing $totalShipments shipments..."
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt $totalShipments; $i++) {
        $shipment = $shipments[$i]; $shipmentNumber = $i + 1
        if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($shipmentNumber * 100 / $totalShipments)) -Message "Processing Shipment $shipmentNumber" }

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'
        $baseCost = Invoke-RLApi -OriginZip $shipment.OriginZip -DestinationZip $shipment.DestinationZip -Weight $shipment.Weight -Class $shipment.Class -KeyData $BaseKeyData -ShipmentDetails $shipment
        $compCost = Invoke-RLApi -OriginZip $shipment.OriginZip -DestinationZip $shipment.DestinationZip -Weight $shipment.Weight -Class $shipment.Class -KeyData $ComparisonKeyData -ShipmentDetails $shipment
        $VerbosePreference = $CurrentVerbosePreference

        if ($baseCost -eq $null -or $compCost -eq $null) { $skippedShipmentCount++; Write-Warning "Skipping shipment $shipmentNumber (API error)."; continue }
        if ($baseCost -le 0) { $skippedShipmentCount++; Write-Warning "Skipping shipment $shipmentNumber (zero/negative Base Cost)."; continue }

        $targetSellPrice = $null; $baseProfit = $null; $compProfit = $null
        try { if ((1.0 - $customerCurrentMarginDecimal) -eq 0) { throw "Division by zero (100% margin)" }; $targetSellPrice = $baseCost / (1.0 - $customerCurrentMarginDecimal); $baseProfit = $targetSellPrice - $baseCost; $compProfit = $targetSellPrice - $compCost }
        catch { Write-Warning "Skipping shipment $shipmentNumber (calculation error: $($_.Exception.Message))"; $skippedShipmentCount++; continue }

        $processedShipmentCount++; $totalTargetSellPrice += $targetSellPrice; $totalCompCost += $compCost
        $resultsData.Add([PSCustomObject]@{ 
            OriginZip = $shipment.OriginZip; DestZip = $shipment.DestinationZip; 
            Weight = $shipment.Weight; Class = $shipment.Class;
            BaseCost = $baseCost; CompCost = $compCost; TargetSellPrice = $targetSellPrice; 
            BaseProfit = $baseProfit; CompProfit = $compProfit 
        })
    }
    if ($useLoadingBar) { Write-Progress -Activity "Processing Shipments" -Completed }

    foreach ($result in $resultsData) {
        try {
             $originZipStr = if ([string]::IsNullOrWhiteSpace($result.OriginZip)) { 'N/A' } else { $result.OriginZip }; $destZipStr = if ([string]::IsNullOrWhiteSpace($result.DestZip)) { 'N/A' } else { $result.DestZip }
             $weightStr = if ($null -eq $result.Weight) { 'N/A' } else { $result.Weight.ToString("N0") }; $classStr = if ([string]::IsNullOrWhiteSpace($result.Class)) { 'N/A' } else { $result.Class }
             $baseCostStr = if ($result.BaseCost -ne $null) { $result.BaseCost.ToString("C2") } else { 'Error' }; $compCostStr = if ($result.CompCost -ne $null) { $result.CompCost.ToString("C2") } else { 'Error' }
             $targetSellPriceStr = if ($result.TargetSellPrice -ne $null) { $result.TargetSellPrice.ToString("C2") } else { 'Error' }; $baseProfitStr = if ($result.BaseProfit -ne $null) { $result.BaseProfit.ToString("C2") } else { 'Error' }
             $compProfitStr = if ($result.CompProfit -ne $null) { $result.CompProfit.ToString("C2") } else { 'Error' }
             $line = ($originZipStr.PadRight($col1Width)) + ($destZipStr.PadRight($col2Width)) + ($weightStr.PadRight($col3Width)) + ($classStr.PadRight($col4Width)) + ($baseCostStr.PadRight($col5Width)) + ($compCostStr.PadRight($col6Width)) + ($targetSellPriceStr.PadRight($col7Width)) + ($baseProfitStr.PadRight($col8Width)) + ($compProfitStr.PadRight($col9Width))
             $reportContent.Add($line)
        } catch { Write-Warning "Skipping result row (Origin: $($result.OriginZip)) due to formatting error: $($_.Exception.Message)"; $skippedShipmentCount++ }
    }

    $reportContent.Add("------------------------------------------------------------------------------------------------------------------------------------"); $reportContent.Add("Summary:"); $reportContent.Add("Processed Shipments (API & Calc OK): $processedShipmentCount"); $reportContent.Add("Skipped Shipments (API/Calc/Formatting Errors): $skippedShipmentCount")
    if ($processedShipmentCount -gt 0) {
        $avgTargetSellPrice = $totalTargetSellPrice / $processedShipmentCount; $avgCompCost = $totalCompCost / $processedShipmentCount
        $reportContent.Add("Average Target Sell Price (from '$baseKeyName' + $customerCurrentMarginPercent%): $($avgTargetSellPrice.ToString("C2"))")
        $reportContent.Add("Average Comparison Cost (from '$compKeyName'): $($avgCompCost.ToString("C2"))")
        $avgRequiredMarginPercent = $null
        if ($avgTargetSellPrice -ne 0) { try { $avgRequiredMarginDecimal = ($avgTargetSellPrice - $avgCompCost) / $avgTargetSellPrice; $avgRequiredMarginPercent = $avgRequiredMarginDecimal * 100.0; $reportContent.Add("Average Margin Required on '$compKeyName' Cost to Match Average Target Sell Price: $($avgRequiredMarginPercent.ToString("N2"))%") } catch { $reportContent.Add("Could not calculate avg required margin.")} }
        else { $reportContent.Add("Average Target Sell Price is zero, cannot calculate average required margin.") }
    } else { $reportContent.Add("No shipments processed successfully, cannot calculate average margin.") }
    $reportContent.Add("------------------------------------------------------------------------------------------------------------------------------------"); $reportContent.Add("End of Report")

    try { $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop; Write-Host "`nAverage Margin Calculation Report saved successfully to: $reportFilePath" -ForegroundColor Green; return $reportFilePath }
    catch { Write-Error "Failed to save report file '$reportFilePath': $($_.Exception.Message)"; return $null }
}


function Calculate-RLMarginForASPReportGUI {
    param(
        [Parameter(Mandatory)][hashtable]$CostAccountInfo, 
        [Parameter(Mandatory)][decimal]$DesiredASP,      
        [Parameter(Mandatory)][string]$CsvFilePath,
        [Parameter(Mandatory)][string]$Username,
        [Parameter(Mandatory)][string]$UserReportsFolder
    )
    Write-Host "`nRunning R+L Required Margin for Desired ASP Report (GUI Mode)..." -ForegroundColor Cyan
    if ($DesiredASP -le 0) { Write-Error "Desired ASP must be positive."; return $null }
    Write-Host "Cost Basis Account: '$($CostAccountInfo.Name)', Desired ASP: $($DesiredASP.ToString('C2'))"

    $shipments = Load-And-Normalize-RLData -CsvPath $CsvFilePath # Use R+L specific normalization
    if ($shipments -eq $null -or $shipments.Count -eq 0) { Write-Warning "No processable shipment data found in '$CsvFilePath'."; return $null }

    $reportContent = [System.Collections.Generic.List[string]]::new(); $resultsData = [System.Collections.Generic.List[object]]::new()
    $safeCostKeyName = $CostAccountInfo.Name -replace '[^a-zA-Z0-9_-]', ''
    $reportFilePath = Get-ReportPath -BaseDir $UserReportsFolder -Username $Username -Carrier 'RL' -ReportType 'MarginForASP' -FilePrefix $safeCostKeyName
    if (-not $reportFilePath) { return $null }

    $skippedShipmentCount = 0; $processedShipmentCount = 0; $totalCostValue = 0.0
    $reportContent.Add("R+L Carriers Required Margin for Desired ASP Report"); $reportContent.Add("User: $Username"); $reportContent.Add("Date: $(Get-Date)"); $reportContent.Add("Data File: $CsvFilePath")
    $reportContent.Add("Cost Basis Account: '$($CostAccountInfo.Name)'"); $reportContent.Add("Desired Average Sell Price (ASP): $($DesiredASP.ToString("C2"))")
    $reportContent.Add("-----------------------------------------------------------------------------------------") 
    $col1ASP = 10; $col2ASP = 10; $col3ASP = 12; $col4ASP = 12; $col5ASP = 15
    $headerLineASP = ("Origin Zip".PadRight($col1ASP)) + ("Dest Zip".PadRight($col2ASP)) + ("Weight".PadRight($col3ASP)) + ("Class".PadRight($col4ASP)) + ("Retrieved Cost".PadRight($col5ASP))
    $reportContent.Add($headerLineASP); $reportContent.Add("-----------------------------------------------------------------------------------------") 

    $totalShipments = $shipments.Count; Write-Host "Processing $totalShipments shipments to get costs..."
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue
    for ($i = 0; $i -lt $totalShipments; $i++) {
        $shipment = $shipments[$i]; $shipmentNumber = $i + 1
        if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($shipmentNumber * 100 / $totalShipments)) -Message "Processing Shipment $shipmentNumber" }

        $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'
        $costValue = Invoke-RLApi -OriginZip $shipment.OriginZip -DestinationZip $shipment.DestinationZip -Weight $shipment.Weight -Class $shipment.Class -KeyData $CostAccountInfo -ShipmentDetails $shipment
        $VerbosePreference = $CurrentVerbosePreference

        if ($costValue -eq $null) { $skippedShipmentCount++; Write-Warning "Skipping shipment $shipmentNumber (API error)."; continue }

        $processedShipmentCount++; $totalCostValue += $costValue
        $resultsData.Add([PSCustomObject]@{ 
            OriginZip = $shipment.OriginZip; DestZip = $shipment.DestinationZip; 
            Weight = $shipment.Weight; Class = $shipment.Class; 
            Cost = $costValue 
        })
    }
    if ($useLoadingBar) { Write-Progress -Activity "Processing Shipments" -Completed }

    foreach ($result in $resultsData) {
        try {
             $originZipStr = if ([string]::IsNullOrWhiteSpace($result.OriginZip)) { 'N/A' } else { $result.OriginZip }; $destZipStr = if ([string]::IsNullOrWhiteSpace($result.DestZip)) { 'N/A' } else { $result.DestZip }
             $weightStr = if ($null -eq $result.Weight) { 'N/A' } else { $result.Weight.ToString("N0") }; $classStr = if ([string]::IsNullOrWhiteSpace($result.Class)) { 'N/A' } else { $result.Class }
             $costStr = if ($result.Cost -ne $null) { $result.Cost.ToString("C2") } else { 'Error' }
             $line = ($originZipStr.PadRight($col1ASP)) + ($destZipStr.PadRight($col2ASP)) + ($weightStr.PadRight($col3ASP)) + ($classStr.PadRight($col4ASP)) + ($costStr.PadRight($col5ASP))
             $reportContent.Add($line)
        } catch { Write-Warning "Skipping result row (Origin: $($result.OriginZip)) due to formatting error: $($_.Exception.Message)"; $skippedShipmentCount++ }
    }

    $reportContent.Add("-----------------------------------------------------------------------------------------"); $reportContent.Add("Summary:"); $reportContent.Add("Processed Shipments (API OK): $processedShipmentCount"); $reportContent.Add("Skipped Shipments (API Errors/Formatting): $skippedShipmentCount")
    $avgCost = 0.0; $requiredAvgMarginPercent = $null; $avgProfitPerShipment = 0.0
    if ($processedShipmentCount -gt 0) {
        $avgCost = $totalCostValue / $processedShipmentCount; $avgProfitPerShipment = $DesiredASP - $avgCost
        if ($DesiredASP -ne 0) { try { $requiredAvgMarginPercent = (($DesiredASP - $avgCost) / $DesiredASP) * 100.0 } catch {} }
        $reportContent.Add("Average Cost (from '$($CostAccountInfo.Name)'): $($avgCost.ToString("C2"))"); $reportContent.Add("Desired Average Sell Price (ASP): $($DesiredASP.ToString("C2"))")
        if ($requiredAvgMarginPercent -ne $null) { $reportContent.Add("Required Avg Margin % on Cost to achieve ASP: $($requiredAvgMarginPercent.ToString("N2"))%") }
        else { $reportContent.Add("Required Avg Margin % on Cost: N/A (Likely due to $0 ASP)") }
        $reportContent.Add("Calculated Avg Profit/Shipment at Desired ASP: $($avgProfitPerShipment.ToString("C2"))")
    } else { $reportContent.Add("No shipments processed successfully, cannot calculate averages.") }
    $reportContent.Add("-----------------------------------------------------------------------------------------"); $reportContent.Add("End of Report")

    try { $reportContent | Out-File -FilePath $reportFilePath -Encoding UTF8 -ErrorAction Stop; Write-Host "`nR+L Carriers Required Margin for ASP Report saved successfully to: $reportFilePath" -ForegroundColor Green; return $reportFilePath }
    catch { Write-Error "Failed to save report file '$reportFilePath': $($_.Exception.Message)"; return $null }
}

Write-Verbose "TMS R+L Carrier Functions loaded (GUI Refactored)."
