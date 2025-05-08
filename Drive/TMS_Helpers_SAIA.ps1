# TMS_Helpers_SAIA.ps1
# Description: Contains helper functions specific to SAIA carrier operations,
#              including data normalization and API interaction.
#              This file should be dot-sourced by the main script(s) after TMS_Config.ps1.

# Assumes config variables like $script:saiaApiUri are available from TMS_Config.ps1
# Assumes general helper functions (if any were used by these) are available.

# --- Data Normalization Functions ---

function Load-And-Normalize-SAIAData {
    # NOTE: This function might need adjustments if your CSV structure changes
    # to support multiple commodities per row for SAIA reports.
    # Currently, it assumes single commodity details per row based on original design.
    param([Parameter(Mandatory)][string]$CsvPath)
    Write-Host "`nLoading SAIA data: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan
    # Base required for API + single commodity mapping from common CSV columns
    $reqCols = @( "Origin Postal Code", "Destination Postal Code", "Total Weight", "Freight Class 1", "Origin City", "Origin State", "Destination City", "Destination State", "Total Units", "Total Length", "Total Width", "Total Height")
    $optCols = @( "Description 1", "Stackable 1" ) # Optional for dims

    try {
        if (-not (Test-Path -Path $CsvPath -PathType Leaf)) { Write-Error "CSV file not found at '$CsvPath'."; return $null }
        $rawData = Import-Csv -Path $CsvPath -ErrorAction Stop
        Write-Host " -> Rows read from CSV: $($rawData.Count)." -ForegroundColor Gray
        if ($rawData.Count -eq 0) { Write-Warning "CSV empty."; return @() }

        $headers = $rawData[0].PSObject.Properties.Name
        $missingReq = $reqCols | Where-Object { $_ -notin $headers }
        if ($missingReq.Count -gt 0) { Write-Error "CSV missing required SAIA columns: $($missingReq -join ', ')"; return $null }
        $missingOpt = $optCols | Where-Object { $_ -notin $headers }
        if($missingOpt.Count -gt 0){ Write-Warning "CSV missing optional SAIA columns: $($missingOpt -join ', ')" }

        $normData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing SAIA data..." -ForegroundColor Gray
        $invalid = 0; $rowNum = 1
        foreach ($row in $rawData) {
            $rowNum++
            # Read values, trimming whitespace
            $oZipRaw=$row."Origin Postal Code"; $dZipRaw=$row."Destination Postal Code"; $wtStrRaw=$row."Total Weight"; $clStrRaw=$row."Freight Class 1"; $oCityRaw=$row."Origin City"; $oStateRaw=$row."Origin State"; $dCityRaw=$row."Destination City"; $dStateRaw=$row."Destination State"; $pcsStrRaw=$row."Total Units"; $lenStrRaw=$row."Total Length"; $widStrRaw=$row."Total Width"; $hgtStrRaw=$row."Total Height"

            $oZip=$oZipRaw.Trim(); $dZip=$dZipRaw.Trim(); $wtStr=$wtStrRaw.Trim(); $clStr=$clStrRaw.Trim(); $oCity=$oCityRaw.Trim(); $oState=$oStateRaw.Trim(); $dCity=$dCityRaw.Trim(); $dState=$dStateRaw.Trim(); $pcsStr=$pcsStrRaw.Trim(); $lenStr=$lenStrRaw.Trim(); $widStr=$widStrRaw.Trim(); $hgtStr=$hgtStrRaw.Trim()
            $wtNum=$null; $pcsNum=$null; $lenNum=$null; $widNum=$null; $hgtNum=$null; $skipRow = $false

            # Basic Validation
            if ([string]::IsNullOrWhiteSpace($oZip) -or $oZip.Length -lt 5) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Origin Zip"; $skipRow = $true }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($dZip) -or $dZip.Length -lt 5)) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Dest Zip"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($clStr)) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Class"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($oCity)) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Origin City"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($oState)) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Origin State"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($dCity)) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Dest City"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($dState)) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Dest State"; $skipRow = $true }
            # Numeric Validation
            if (-not $skipRow) { try { $wtNum = [decimal]$wtStr; if($wtNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Weight"; $skipRow = $true } }
            if (-not $skipRow) { try { $pcsNum = [int]$pcsStr; if($pcsNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Pieces"; $skipRow = $true } }
            if (-not $skipRow) { try { $lenNum = [decimal]$lenStr; if($lenNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Length"; $skipRow = $true } }
            if (-not $skipRow) { try { $widNum = [decimal]$widStr; if($widNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Width"; $skipRow = $true } }
            if (-not $skipRow) { try { $hgtNum = [decimal]$hgtStr; if($hgtNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Height"; $skipRow = $true } }

            if ($skipRow) { continue }

            # Create a single commodity item for reports
            $commodityItem = [ordered]@{
                classification = $clStr
                weight = $wtNum # Keep as number for potential calculations later
                length = $lenNum
                width = $widNum
                height = $hgtNum
                pieces = $pcsNum
                packagingType = if ($headers -contains 'Packaging Type 1') { $row.'Packaging Type 1'.Trim() } else { "PAT" }
                description = if ($headers -contains 'Description 1') { $row.'Description 1'.Trim() } else { "Commodity" }
                stackable = if ($headers -contains 'Stackable 1') { $row.'Stackable 1'.Trim() } else { "Y" }
            }

            $newRow = [PSCustomObject]@{
                OriginZip        = $oZip
                DestinationZip   = $dZip
                OriginCity       = $oCity
                OriginState      = $oState
                DestinationCity  = $dCity
                DestinationState = $dState
                # Store as array for consistency with Invoke-SAIAApi multi-item format
                details          = @($commodityItem)
                # Include original total weight/class for potential historical lookups if needed
                'Total Weight'   = $wtNum
                'Freight Class 1' = $clStr
            }
            $normData.Add($newRow)
        }

        if ($invalid -gt 0) { Write-Warning " -> Skipped $invalid SAIA rows during normalization (missing/invalid essential data)." }
        Write-Host " -> OK: $($normData.Count) SAIA rows normalized." -ForegroundColor Green
        return $normData
    } catch {
        Write-Error "Error processing SAIA CSV '$CsvPath': $($_.Exception.Message)"; return $null
    }
}

# --- API Call Functions ---

function Invoke-SAIAApi {
    param(
        [Parameter(Mandatory=$true)] [hashtable]$KeyData,
        [Parameter(Mandatory=$true)] [string]$OriginZip,
        [Parameter(Mandatory=$true)] [string]$DestinationZip,
        [Parameter(Mandatory=$true)] [string]$OriginCity,
        [Parameter(Mandatory=$true)] [string]$OriginState,
        [Parameter(Mandatory=$true)] [string]$DestinationCity,
        [Parameter(Mandatory=$true)] [string]$DestinationState,
        # <<< PARAMETER CHANGE: Accept array of commodities via $Details >>>
        [Parameter(Mandatory=$true)] [array]$Details # Array of hashtables/PSObjects for commodities
        # Removed single Weight/Class params
    )

    $saiaUserID = $null; $saiaPassword = $null; $saiaRQKey = $null; $accountCodeToUse = $null
    $tariffNameForLog = if ($KeyData.ContainsKey('TariffFileName')) { $KeyData.TariffFileName } elseif ($KeyData.ContainsKey('Name')) {$KeyData.Name} else { "UnknownSAIATariff" }

    # Extract credentials from KeyData
    try {
        if ($KeyData.ContainsKey('UserID')) { $saiaUserID = $KeyData.UserID } else { Write-Verbose "UserID missing from KeyData for Acct '$tariffNameForLog'."}
        if ($KeyData.ContainsKey('Password')) { $saiaPassword = $KeyData.Password } else { Write-Verbose "Password missing from KeyData for Acct '$tariffNameForLog'."}
        if ($KeyData.ContainsKey('RQKey')) { $saiaRQKey = $KeyData.RQKey } else { Write-Verbose "RQKey missing from KeyData for Acct '$tariffNameForLog'."}
        if ($KeyData.ContainsKey('AccountCode')) { $accountCodeToUse = $KeyData.AccountCode } else { Write-Verbose "AccountCode not found in KeyData for SAIA Acct '$tariffNameForLog'." }
    } catch { Write-Warning "SAIA Credential Extract Fail for Acct '$tariffNameForLog': $($_.Exception.Message)"; return $null }

    # --- Validate Inputs ---
    $missingFields = @()
    if ([string]::IsNullOrWhiteSpace($OriginZip)) { $missingFields += "OriginZip" }
    if ([string]::IsNullOrWhiteSpace($DestinationZip)) { $missingFields += "DestinationZip" }
    if ([string]::IsNullOrWhiteSpace($OriginCity)) { $missingFields += "OriginCity" }
    if ([string]::IsNullOrWhiteSpace($OriginState)) { $missingFields += "OriginState" }
    if ([string]::IsNullOrWhiteSpace($DestinationCity)) { $missingFields += "DestinationCity" }
    if ([string]::IsNullOrWhiteSpace($DestinationState)) { $missingFields += "DestinationState" }
    if ($null -eq $Details -or $Details.Count -eq 0) { $missingFields += "Details (Commodity array empty or null)"}
    if ([string]::IsNullOrWhiteSpace($saiaRQKey) -and ([string]::IsNullOrWhiteSpace($saiaUserID) -or [string]::IsNullOrWhiteSpace($saiaPassword))) { $missingFields += "Credentials (Missing RQKey AND UserID/Password pair)" }

    # --- Process Commodity Details Array ---
    $detailsArrayForPayload = @()
    $totalWeightForPayload = 0.0
    $totalCubeInches = 0.0
    if ($Details -is [array] -and $Details.Count -gt 0) {
        for ($i = 0; $i -lt $Details.Count; $i++) {
            $item = $Details[$i]
            $itemWeight = $null; $itemClass = $null; $itemLength = $null; $itemWidth = $null; $itemHeight = $null; $itemPieces = $null; $itemStackable = 'N' # Default stackable to No
            $isValidItem = $true

            if ($item -is [hashtable] -or $item -is [psobject]) {
                # Extract and validate each field for the current item
                if ($item.PSObject.Properties.Name -contains 'weight' -and $item.weight -as [decimal] -ne $null -and [decimal]$item.weight -gt 0) { $itemWeight = [decimal]$item.weight } else { $isValidItem = $false; $missingFields += "Item $($i+1) Invalid Weight '$($item.weight)'" }
                if ($item.PSObject.Properties.Name -contains 'classification' -and -not [string]::IsNullOrWhiteSpace($item.classification)) { $itemClass = $item.classification } elseif ($item.PSObject.Properties.Name -contains 'class' -and -not [string]::IsNullOrWhiteSpace($item.class)) { $itemClass = $item.class } else { $isValidItem = $false; $missingFields += "Item $($i+1) Invalid Class" } # Accept 'class' or 'classification'
                if ($item.PSObject.Properties.Name -contains 'length' -and $item.length -as [decimal] -ne $null -and [decimal]$item.length -gt 0) { $itemLength = [decimal]$item.length } else { $isValidItem = $false; $missingFields += "Item $($i+1) Invalid Length '$($item.length)'" }
                if ($item.PSObject.Properties.Name -contains 'width' -and $item.width -as [decimal] -ne $null -and [decimal]$item.width -gt 0) { $itemWidth = [decimal]$item.width } else { $isValidItem = $false; $missingFields += "Item $($i+1) Invalid Width '$($item.width)'" }
                if ($item.PSObject.Properties.Name -contains 'height' -and $item.height -as [decimal] -ne $null -and [decimal]$item.height -gt 0) { $itemHeight = [decimal]$item.height } else { $isValidItem = $false; $missingFields += "Item $($i+1) Invalid Height '$($item.height)'" }
                if ($item.PSObject.Properties.Name -contains 'pieces' -and $item.pieces -as [int] -ne $null -and [int]$item.pieces -gt 0) { $itemPieces = [int]$item.pieces } else { $isValidItem = $false; $missingFields += "Item $($i+1) Invalid Pieces '$($item.pieces)'" }
                if ($item.PSObject.Properties.Name -contains 'stackable' -and $item.stackable -eq 'Y') {$itemStackable = 'Y'} # Only set to Y if explicitly Y

                if ($isValidItem) {
                    $detailsArrayForPayload += [ordered]@{
                        length = [double]$itemLength # API expects number
                        width  = [double]$itemWidth
                        height = [double]$itemHeight
                        weight = [int]$itemWeight # API example shows integer weight
                        class  = [double]$itemClass # API example shows number class
                        units  = $itemPieces # API uses 'units'
                        # packagingType and description seem optional for rating based on example
                    }
                    $totalWeightForPayload += $itemWeight
                    $totalCubeInches += ($itemLength * $itemWidth * $itemHeight * $itemPieces)
                }
            } else {
                 $missingFields += "Item $($i+1) is not a valid object."
            }
        }
        if ($detailsArrayForPayload.Count -eq 0 -and $Details.Count -gt 0) {
             $missingFields += "No valid commodity items found after validation."
        }
    }
    # End Commodity Processing

    if ($missingFields.Count -gt 0) {
        $contextZips = "OZip:$OriginZip DZip:$DestinationZip"
        Write-Warning "SAIA Skip: Acct '$tariffNameForLog' $contextZips - Missing/Invalid required data: $($missingFields -join ', ')."
        return $null
    }

    # --- Calculate total cube in CUFT ---
    $calculatedTotalCube = 0.0
    if ($totalCubeInches -gt 0) {
         try { $calculatedTotalCube = [Math]::Round($totalCubeInches / 1728, 2) }
         catch { Write-Warning "Could not calculate totalCube for Acct '$tariffNameForLog': $($_.Exception.Message)" }
    }

    # --- Construct Payload ---
    $payloadObject = [ordered]@{
        userID = if([string]::IsNullOrWhiteSpace($saiaUserID)) {$null} else {$saiaUserID} # Omit if blank
        password = if([string]::IsNullOrWhiteSpace($saiaPassword)) {$null} else {$saiaPassword} # Omit if blank
        payer = "Shipper" # Assuming Shipper, adjust if needed based on KeyData or other logic
        pickUpDate = (Get-Date -Format 'yyyy-MM-dd')
        origin = @{
            city = $OriginCity
            state = $OriginState
            zipcode = $OriginZip
            # accountCode = $accountCodeToUse # Add if applicable
        }
        destination = @{
            city = $DestinationCity
            state = $DestinationState
            zipcode = $DestinationZip
            # accountCode = $accountCodeToUse # Add if applicable
        }
        weightUnits = "LBS"
        measurementUnit = "IN"
        totalCube = $calculatedTotalCube
        totalCubeUnits = "CUFT"
        details = $detailsArrayForPayload # Use the processed array
        # Accessorials omitted for simplicity, add if needed
    }

    # Conditionally add accountCode if provided
    if (-not [string]::IsNullOrWhiteSpace($accountCodeToUse)) {
        $payloadObject.origin.accountCode = $accountCodeToUse
        $payloadObject.destination.accountCode = $accountCodeToUse
    }

    # Remove null credentials if RQKey is used
    if (-not [string]::IsNullOrWhiteSpace($saiaRQKey)) {
        $payloadObject.Remove('userID')
        $payloadObject.Remove('password')
    }

    $payload = $payloadObject | ConvertTo-Json -Depth 10
    Write-Verbose "SAIA Payload ($tariffNameForLog): $payload"

    # --- API Call ---
    $headers = @{ 'Content-Type' = 'application/json'; 'Cache-Control' = 'no-cache' }
    if (-not [string]::IsNullOrWhiteSpace($saiaRQKey)) {
        $headers.'RQ-Key' = $saiaRQKey
    }

    try {
        $apiUrl = $script:saiaApiUri
        if ([string]::IsNullOrWhiteSpace($apiUrl)) { throw "SAIA API URI ($($apiUrl)) is not defined or empty."}

        # Final credential check before call
        if (-not $headers.ContainsKey('RQ-Key') -and ([string]::IsNullOrWhiteSpace($payloadObject.userID) -or [string]::IsNullOrWhiteSpace($payloadObject.password)) ) {
             throw "Cannot call SAIA API for Acct '$tariffNameForLog': No RQ-Key header provided AND UserID/Password pair is incomplete/missing in payload."
        }

        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $payload -ErrorAction Stop
        Write-Verbose "SAIA OK: Acct '$tariffNameForLog'"

        $totalChargeValue = $null
        if ($response -ne $null -and
            $response.PSObject.Properties.Name -contains 'rateDetails' -and
            $response.rateDetails -ne $null -and
            $response.rateDetails.PSObject.Properties.Name -contains 'totalInvoice') {
             $totalChargeValue = $response.rateDetails.totalInvoice
        }

        if ($totalChargeValue -ne $null) {
            try {
                 $cleanedRate = $totalChargeValue -replace '[$,]'
                 $decimalRate = [decimal]$cleanedRate
                 if ($decimalRate -lt 0) { throw "Negative rate returned."}
                 return $decimalRate
            } catch {
                 Write-Warning "SAIA Convert Fail for Acct '$tariffNameForLog': Cannot convert rate '$totalChargeValue' to decimal. Error: $($_.Exception.Message)"; return $null
            }
        } else {
             Write-Warning "SAIA Resp missing 'rateDetails.totalInvoice' or structure invalid for Acct '$tariffNameForLog'.";
             # Write-Verbose "SAIA Full Response: $($response | ConvertTo-Json -Depth 5)" # Uncomment for deep debug
             return $null
        }
    } catch {
        $errMsg = $_.Exception.Message; $statusCode = "N/A"; $eBody = "N/A"
        if ($_.Exception.Response) {
             try {$statusCode = $_.Exception.Response.StatusCode.value__} catch{}
             try { $stream = $_.Exception.Response.GetResponseStream(); $reader = New-Object System.IO.StreamReader($stream); $eBody = $reader.ReadToEnd(); $reader.Close(); $stream.Close() } catch {$eBody="(Err reading resp body: $($_.Exception.Message))"}
        }
        $truncatedBody = if ($eBody.Length -gt 500) { $eBody.Substring(0, 500) + "..." } else { $eBody }
        $fullErrMsg = "SAIA FAIL: Acct '$tariffNameForLog'. Error: $errMsg (HTTP $statusCode) Resp: $truncatedBody"
        Write-Warning $fullErrMsg;
        return $null
     }
}

Write-Verbose "TMS SAIA Helper Functions loaded."