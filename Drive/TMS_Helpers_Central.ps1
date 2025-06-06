# TMS_Helpers_Central.ps1
# Description: Contains helper functions specific to Central Transport carrier operations,
#              including data normalization and API interaction.
#              This file should be dot-sourced by the main script(s) after TMS_Config.ps1.

# Assumes config variables like $script:centralApiUri are available from TMS_Config.ps1
# Assumes general helper functions (if any were used by these) are available.

# --- Data Normalization Functions ---

function Load-And-Normalize-CentralData {
    # NOTE: This function might need adjustments if your CSV structure changes
    # to support multiple commodities per row for Central Transport reports.
    # Currently, it assumes single commodity details per row based on original design.
    param([Parameter(Mandatory)][string]$CsvPath)
    Write-Host "`nLoading Central data: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan
    # Required for single item mapping
    $reqCols = @("Origin Postal Code", "Destination Postal Code", "Total Weight", "Freight Class 1", "Total Units", "Total Length", "Total Width", "Total Height")
    try {
        if (-not (Test-Path -Path $CsvPath -PathType Leaf)) {
            Write-Error "CSV file not found at '$CsvPath'."
            return $null
        }
        $rawData = Import-Csv -Path $CsvPath -ErrorAction Stop
        Write-Host " -> Rows read from CSV: $($rawData.Count)." -ForegroundColor Gray
        if ($rawData.Count -eq 0) { Write-Warning "CSV empty."; return @() } # Return empty array for no data

        $headers = $rawData[0].PSObject.Properties.Name
        $missing = $reqCols | Where-Object { $_ -notin $headers }
        if ($missing.Count -gt 0) { Write-Error "CSV missing Central columns: $($missing -join ', ')"; return $null }

        $normData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing Central data..." -ForegroundColor Gray
        $invalid = 0; $rowNum = 1
        foreach ($row in $rawData) {
            $rowNum++
            $oZipRaw=$row."Origin Postal Code"; $dZipRaw=$row."Destination Postal Code"; $wtStrRaw=$row."Total Weight"; $clStrRaw=$row."Freight Class 1"; $pcsStrRaw=$row."Total Units"; $lenStrRaw=$row."Total Length"; $widStrRaw=$row."Total Width"; $hgtStrRaw=$row."Total Height"
            $oZip=$oZipRaw.Trim(); $dZip=$dZipRaw.Trim(); $wtStr=$wtStrRaw.Trim(); $clStr=$clStrRaw.Trim(); $pcsStr=$pcsStrRaw.Trim(); $lenStr=$lenStrRaw.Trim(); $widStr=$widStrRaw.Trim(); $hgtStr=$hgtStrRaw.Trim()
            $wtNum=$null; $pcsNum=$null; $lenNum=$null; $widNum=$null; $hgtNum=$null;
            $skipRow = $false

            if ([string]::IsNullOrWhiteSpace($oZip) -or $oZip.Length -lt 5) { $invalid++; Write-Verbose "Skip CTX Row ${rowNum}: Bad Origin Zip"; $skipRow = $true }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($dZip) -or $dZip.Length -lt 5)) { $invalid++; Write-Verbose "Skip CTX Row ${rowNum}: Bad Dest Zip"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($clStr)) { $invalid++; Write-Verbose "Skip CTX Row ${rowNum}: Bad Class"; $skipRow = $true } # Class not needed for API, but good for data integrity check
            # Numeric Validation
            if (-not $skipRow) { try { $wtNum = [decimal]$wtStr; if($wtNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip CTX Row ${rowNum}: Bad Weight"; $skipRow = $true } }
            if (-not $skipRow) { try { $pcsNum = [int]$pcsStr; if($pcsNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip CTX Row ${rowNum}: Bad Pieces"; $skipRow = $true } }
            if (-not $skipRow) { try { $lenNum = [decimal]$lenStr; if($lenNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip CTX Row ${rowNum}: Bad Length"; $skipRow = $true } }
            if (-not $skipRow) { try { $widNum = [decimal]$widStr; if($widNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip CTX Row ${rowNum}: Bad Width"; $skipRow = $true } }
            if (-not $skipRow) { try { $hgtNum = [decimal]$hgtStr; if($hgtNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip CTX Row ${rowNum}: Bad Height"; $skipRow = $true } }


            if (-not $skipRow) {
                 # Create a single commodity item reflecting the API structure
                 $weightPerUnit = 0.0
                 if ($pcsNum -gt 0) { $weightPerUnit = [math]::Round($wtNum / $pcsNum, 2) }

                 $commodityItem = [ordered]@{
                    id = 1 # Central API uses 'id' for rateItems
                    handlingUnits = $pcsNum
                    weightPerHandlingUnit = $weightPerUnit
                    width = $widNum
                    length = $lenNum
                    height = $hgtNum
                    # Store original values from CSV if needed for other purposes
                    pieces = $pcsNum # Example - Store original pieces if needed later
                    weight = $wtNum # Example - Store original total weight if needed later
                    itemClass = $clStr # Example - Store original class if needed later
                 }
                 $normData.Add([PSCustomObject]@{
                    "Origin Postal Code" = $oZip
                    "Destination Postal Code" = $dZip
                    "Commodities" = @($commodityItem) # Store as an array
                    "Freight Class 1" = $clStr # Keep original class info if needed
                 })
            }
        }
        if ($invalid -gt 0) { Write-Warning " -> Skipped $invalid Central rows (missing/invalid data)." }
        Write-Host " -> OK: $($normData.Count) Central rows normalized." -ForegroundColor Green
        return $normData
    } catch {
        Write-Error "Error processing Central CSV '$CsvPath': $($_.Exception.Message)"
        return $null
    }
}

# --- API Call Functions ---

function Invoke-CentralTransportApi {
    [CmdletBinding(DefaultParameterSetName = 'FromShipmentObject')]
    param(
        # Parameter Set 'FromShipmentObject' - Used by Reports
        [Parameter(ParameterSetName='FromShipmentObject')]
        [PSCustomObject]$ShipmentData, # Contains Origin/Dest Zips and Commodities array

        [Parameter(ParameterSetName='FromShipmentObject')]
        [hashtable]$KeyData,

        # Parameter Set 'FromIndividualParams' - Used by Quote Tab
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')]
        [hashtable]$ApiKeyData, # Contains accessCode, customerNumber
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')]
        [string]$OriginZip,
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')]
        [string]$DestinationZip,
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')]
        [array]$Commodities, # Array of hashtables/PSObjects from GUI Grid
        [Parameter(ParameterSetName='FromIndividualParams')]
        [string]$Accessorials = $null # Placeholder
    )

    $accessCodeToUse = $null; $customerNumberToUse = $null; $originZipToUse = $null; $destZipToUse = $null; $commoditiesToUse = $null; $tariffNameForLog = "Unknown"
    $localKeyData = $null

    # Determine parameters based on Parameter Set
    if ($PSCmdlet.ParameterSetName -eq 'FromShipmentObject') {
        # Logic for reports (using normalized data from Load-And-Normalize-CentralData)
        $localKeyData = $KeyData
        try {
            if (!$ShipmentData) { throw "ShipmentData parameter is null." }
            if (!$ShipmentData.'Origin Postal Code') { throw "ShipmentData missing 'Origin Postal Code'."}
            if (!$ShipmentData.'Destination Postal Code') { throw "ShipmentData missing 'Destination Postal Code'."}
            if (!$ShipmentData.Commodities) { throw "ShipmentData missing 'Commodities' array."}
            $originZipToUse = $ShipmentData.'Origin Postal Code'
            $destZipToUse = $ShipmentData.'Destination Postal Code'
            $commoditiesToUse = $ShipmentData.Commodities # Already contains the required structure from normalization
            if (!$localKeyData) { throw "KeyData parameter is null." }
            if (!$localKeyData.accessCode) { throw "'accessCode' missing from KeyData." }
            if (!$localKeyData.customerNumber) { throw "'customerNumber' missing from KeyData." }
            $accessCodeToUse = $localKeyData.accessCode; $customerNumberToUse = $localKeyData.customerNumber
            $tariffNameForLog = if ($localKeyData.TariffFileName) { $localKeyData.TariffFileName } elseif($localKeyData.Name) {$localKeyData.Name} else { "UnknownTariff" }
        } catch { Write-Warning "Central Extract Fail (Report): $($_.Exception.Message)"; return $null }
    } elseif ($PSCmdlet.ParameterSetName -eq 'FromIndividualParams') {
         # Logic for Quote Tab
         $localKeyData = $ApiKeyData
         if (!$localKeyData) { Write-Warning "ApiKeyData parameter is null."; return $null }
         if (!$localKeyData.accessCode) { Write-Warning "AccessCode missing from ApiKeyData."; return $null }
         if (!$localKeyData.customerNumber) { Write-Warning "CustomerNumber missing from ApiKeyData."; return $null }
         $accessCodeToUse = $localKeyData.accessCode; $customerNumberToUse = $localKeyData.customerNumber
         $tariffNameForLog = if ($localKeyData.TariffFileName) { $localKeyData.TariffFileName } elseif($localKeyData.Name) {$localKeyData.Name} else { "SingleQuoteCall" }
         $originZipToUse = $OriginZip; $destZipToUse = $DestinationZip; $commoditiesToUse = $Commodities
    } else { Write-Error "CTX Internal Error: Invalid parameter set."; return $null }

    # --- Validate Inputs ---
    $ctxMissing = @()
    if ([string]::IsNullOrWhiteSpace($originZipToUse)) { $ctxMissing += "OriginZip" }
    if ([string]::IsNullOrWhiteSpace($destZipToUse)) { $ctxMissing += "DestinationZip" }
    if ($null -eq $commoditiesToUse -or $commoditiesToUse.Count -eq 0) { $ctxMissing += "Commodities (empty or null)" }
    if ([string]::IsNullOrWhiteSpace($accessCodeToUse)) { $ctxMissing += "AccessCode(ApiKey)" }
    if ([string]::IsNullOrWhiteSpace($customerNumberToUse)) { $ctxMissing += "CustomerNumber" }

    # --- Validate and Format Commodities for API Payload ---
    $validCommoditiesForPayload = @()
    if ($commoditiesToUse -is [array]) {
        for ($i = 0; $i -lt $commoditiesToUse.Count; $i++) {
            $item = $commoditiesToUse[$i]
            $itemPieces = $null; $itemWeight = $null; $itemLength = $null; $itemWidth = $null; $itemHeight = $null;
            $isValidItem = $true

            # 1. Check if item itself is a valid object (not null)
            if ($null -eq $item) {
                $isValidItem = $false; $ctxMissing += "Item $($i+1) is null."; continue # Skip to next item
            }

            # 2. Attempt to access properties directly and validate values
            try {
                $itemPieces = $item.pieces
                if (($itemPieces -as [int]) -eq $null -or ([int]$itemPieces -le 0)) { $isValidItem = $false; $ctxMissing += "Item $($i+1) invalid pieces value '$($itemPieces)'" }

                $itemWeight = $item.weight
                if (($itemWeight -as [decimal]) -eq $null -or ([decimal]$itemWeight -le 0)) { $isValidItem = $false; $ctxMissing += "Item $($i+1) invalid weight value '$($itemWeight)'" }

                $itemLength = $item.length
                if (($itemLength -as [decimal]) -eq $null -or ([decimal]$itemLength -le 0)) { $isValidItem = $false; $ctxMissing += "Item $($i+1) invalid length value '$($itemLength)'" }

                $itemWidth = $item.width
                if (($itemWidth -as [decimal]) -eq $null -or ([decimal]$itemWidth -le 0)) { $isValidItem = $false; $ctxMissing += "Item $($i+1) invalid width value '$($itemWidth)'" }

                $itemHeight = $item.height
                if (($itemHeight -as [decimal]) -eq $null -or ([decimal]$itemHeight -le 0)) { $isValidItem = $false; $ctxMissing += "Item $($i+1) invalid height value '$($itemHeight)'" }

            } catch {
                # Catch errors if properties don't exist on $item
                $isValidItem = $false; $ctxMissing += "Item $($i+1) missing required properties (pieces/weight/dims) or access error: $($_.Exception.Message)"
            }

            # 3. If still valid, calculate weight per unit and add to payload list
            if ($isValidItem) {
                $weightPerUnit = 0.0
                try {
                    if ([int]$itemPieces -eq 0) { throw "Cannot divide by zero pieces."}
                    $weightPerUnit = [math]::Round(([decimal]$itemWeight / [int]$itemPieces), 2)
                    if ($weightPerUnit -lt 0) { throw "Calculated negative weight per unit." }
                } catch {
                     $isValidItem = $false; $ctxMissing += "Item $($i+1) calculation error for weight/unit: $($_.Exception.Message)"
                }

                if ($isValidItem) {
                    # Add item using the keys expected by Central Transport API
                    $validCommoditiesForPayload += [ordered]@{
                        id = $i + 1
                        handlingUnits = [int]$itemPieces
                        weightPerHandlingUnit = $weightPerUnit # Use calculated value
                        width = [decimal]$itemWidth # API expects number
                        length = [decimal]$itemLength
                        height = [decimal]$itemHeight
                    }
                }
            }
        } # End for loop

        if ($validCommoditiesForPayload.Count -eq 0 -and $commoditiesToUse.Count -gt 0) {
             $ctxMissing += "No valid commodity items found after validation."
        }
    } else {
         $ctxMissing += "Commodities parameter is not an array."
    }


    if ($ctxMissing.Count -gt 0) {
        $logTariffName = if($localKeyData?.TariffFileName){$localKeyData.TariffFileName}elseif($localKeyData?.Name){$localKeyData.Name}else{$tariffNameForLog}
        Write-Warning "CTX Skip: Tariff '$($logTariffName)' - Missing/Invalid required data: $($ctxMissing -join ', ')."
        return $null
    }

    # --- Construct Payload with Correct rateItems Structure ---
    $payload = @{
        accessCode = $accessCodeToUse;
        request = @{
            originZipCode = $originZipToUse;
            destinationZipCode = $destZipToUse;
            customerNumber = [string]$customerNumberToUse;
            pickupDate = (Get-Date -Format 'MM/dd/yyyy');
            customerRole = "shipper"; # Assuming shipper for now
            rateItems = $validCommoditiesForPayload; # Use the correctly structured array
            useDefaultTariff = $false
        }
    } | ConvertTo-Json -Depth 5

    # --- API Call ---
    $headers = @{ 'Content-Type' = 'application/json' }
    $logTariffNameCall = if($localKeyData?.TariffFileName){$localKeyData.TariffFileName}elseif($localKeyData?.Name){$localKeyData.Name}else{$tariffNameForLog}
    Write-Verbose "Calling CTX API: Tariff $($logTariffNameCall)"
    # Write-Host "DEBUG: CTX Payload: $payload" # Uncomment for deep payload debugging

    try {
        $apiUrl = $script:centralApiUri
        if ([string]::IsNullOrWhiteSpace($apiUrl)) { throw "Central API URI not configured."}
        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $payload -ErrorAction Stop
        Write-Verbose "CTX OK: Tariff $($logTariffNameCall)"
        $totalChargeValue = $null
        if ($response -ne $null -and $response.PSObject.Properties.Name -contains 'rateTotal') {
            $totalChargeValue = $response.rateTotal
        }
        if ($totalChargeValue -ne $null) {
            try {
                 $cleanedRate = $totalChargeValue -replace '[$,]'; $decimalRate = [decimal]$cleanedRate
                 if ($decimalRate -lt 0) { throw "Negative rate." }
                 return $decimalRate
            } catch { Write-Warning "CTX Convert Fail ($($logTariffNameCall)): Cannot convert '$totalChargeValue'. Error: $($_.Exception.Message)"; return $null }
        } else { Write-Warning "CTX Resp missing 'rateTotal': Tariff $($logTariffNameCall)"; return $null }
    } catch {
         $errMsg = $_.Exception.Message; $statusCode = "N/A"; $eBody = "N/A"
         if ($_.Exception.Response) { try {$statusCode = $_.Exception.Response.StatusCode.value__} catch{}; try { $stream = $_.Exception.Response.GetResponseStream(); $reader = New-Object System.IO.StreamReader($stream); $eBody = $reader.ReadToEnd(); $reader.Close(); $stream.Close() } catch {$eBody="(Err reading resp body)"} }
         $truncatedBody = if ($eBody.Length -gt 500) { $eBody.Substring(0, 500) + "..." } else { $eBody }
         $fullErrMsg = "CTX FAIL: Tariff $($logTariffNameCall). Error: $errMsg (HTTP $statusCode) Resp: $truncatedBody"; Write-Warning $fullErrMsg; return $null
    }
}

Write-Verbose "TMS Central Transport Helper Functions loaded."