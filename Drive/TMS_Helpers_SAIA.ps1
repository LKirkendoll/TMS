# TMS_Helpers_SAIA.ps1
# Description: Contains helper functions specific to SAIA carrier operations,
#              including data normalization and API interaction. Ensure PSCustomObject for commodity item.
#              This file should be dot-sourced by the main script(s) after TMS_Config.ps1.

# Assumes config variables like $script:saiaApiUri are available from TMS_Config.ps1
# Assumes general helper functions (if any were used by these) are available.

# --- Data Normalization Functions ---

function Load-And-Normalize-SAIAData {
    param(
        [Parameter(Mandatory=$true)][string]$CsvPath
    )
    Write-Host "`nLoading SAIA data: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan
    # Base required for API + single commodity mapping from common CSV columns
    $reqCols = @( "Origin Postal Code", "Destination Postal Code", "Total Weight", "Freight Class 1", "Origin City", "Origin State", "Destination City", "Destination State", "Total Units", "Total Length", "Total Width", "Total Height")
    $optCols = @( "Description 1", "Stackable 1" ) # Optional for dims

    try {
        if (-not (Test-Path -Path $CsvPath -PathType Leaf)) { Write-Error "CSV file not found at '$CsvPath'."; return $null }

        # Use Import-Csv for the whole file, then get header (More robust than Get-Content/ConvertFrom-Csv for header)
        $rawData = $null
        try {
            # Explicitly specify delimiter for robustness
            $rawData = Import-Csv -Path $CsvPath -Delimiter ',' -ErrorAction Stop
        } catch {
            Write-Error "Failed to import CSV '$CsvPath'. Error: $($_.Exception.Message)"
            return $null
        }

        if ($null -eq $rawData -or $rawData.Count -eq 0) {
            Write-Warning "CSV '$CsvPath' is empty or could not be imported properly."
            return @() # Return empty array if no data
        }

        # Get header from the first imported object's properties
        $header = $rawData[0].PSObject.Properties.Name
        if ($null -eq $header -or $header.Count -eq 0) {
            Write-Error "Could not extract header information after importing CSV '$CsvPath'."
            return $null
        }

        Write-Host " -> Rows read from CSV: $($rawData.Count)." -ForegroundColor Gray

        $missingReq = $reqCols | Where-Object { $_ -notin $header }
        if ($missingReq.Count -gt 0) { Write-Error "CSV missing required SAIA columns: $($missingReq -join ', ')"; return $null }
        $missingOpt = $optCols | Where-Object { $_ -notin $header }
        if($missingOpt.Count -gt 0){ Write-Warning "CSV missing optional SAIA columns: $($missingOpt -join ', ')" }

        $normData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing SAIA data..." -ForegroundColor Gray
        $invalid = 0; $rowNum = 1 # Start rowNum at 1 for messages (CSV header is effectively row 0 for data rows)
        foreach ($row in $rawData) {
            # For messages, $rowNum will be 2 for the first data row, 3 for the second, etc.
            # This aligns with typical spreadsheet row numbering if header is row 1.
            $currentDataRowForMessage = $rowNum + 1

            # Read values, trimming whitespace
            $oZipRaw=$row."Origin Postal Code"; $dZipRaw=$row."Destination Postal Code"; $wtStrRaw=$row."Total Weight";
            # DEBUG: Check the raw value read for Freight Class 1
            $clStrRaw = $null # Initialize to null
            # Check if the property 'Freight Class 1' exists on the $row object
            if ($row.PSObject.Properties['Freight Class 1'] -ne $null) {
                $clStrRaw = $row."Freight Class 1" # Access the property value
                Write-Host "DEBUG (Load-And-Normalize-SAIAData): DataRow ${currentDataRowForMessage} - Raw 'Freight Class 1' value = [$clStrRaw]" # DEBUG Line
            } else {
                 Write-Host "DEBUG (Load-And-Normalize-SAIAData): DataRow ${currentDataRowForMessage} - Column 'Freight Class 1' not found on row object or its value is \$null." # DEBUG Line
            }

            $oCityRaw=$row."Origin City"; $oStateRaw=$row."Origin State"; $dCityRaw=$row."Destination City"; $dStateRaw=$row."Destination State"; $pcsStrRaw=$row."Total Units"; $lenStrRaw=$row."Total Length"; $widStrRaw=$row."Total Width"; $hgtStrRaw=$row."Total Height"

            $oZip=$oZipRaw.Trim(); $dZip=$dZipRaw.Trim(); $wtStr=$wtStrRaw.Trim();
            # Trim $clStrRaw only if it's not null
            $clStr = if ($clStrRaw -ne $null) { $clStrRaw.Trim() } else { $clStrRaw }
            $oCity=$oCityRaw.Trim(); $oState=$oStateRaw.Trim(); $dCity=$dCityRaw.Trim(); $dState=$dStateRaw.Trim(); $pcsStr=$pcsStrRaw.Trim(); $lenStr=$lenStrRaw.Trim(); $widStr=$widStrRaw.Trim(); $hgtStr=$hgtStrRaw.Trim()
            Write-Host "DEBUG (Load-And-Normalize-SAIAData): DataRow ${currentDataRowForMessage} - Trimmed 'Freight Class 1' value for \$clStr = [$clStr]" # DEBUG Line

            $wtNum=$null; $pcsNum=$null; $lenNum=$null; $widNum=$null; $hgtNum=$null; $skipRow = $false

            # Basic Validation
            if ([string]::IsNullOrWhiteSpace($oZip) -or $oZip.Length -lt 5) { $invalid++; Write-Host "Skip SAIA DataRow ${currentDataRowForMessage}: Bad Origin Zip '$($oZipRaw)'"; $skipRow = $true }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($dZip) -or $dZip.Length -lt 5)) { $invalid++; Write-Host "Skip SAIA DataRow ${currentDataRowForMessage}: Bad Dest Zip '$($dZipRaw)'"; $skipRow = $true }
            # Class validation happens later in Invoke-SAIAApi based on the normalized data
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($oCity)) { $invalid++; Write-Host "Skip SAIA DataRow ${currentDataRowForMessage}: Bad Origin City '$($oCityRaw)'"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($oState)) { $invalid++; Write-Host "Skip SAIA DataRow ${currentDataRowForMessage}: Bad Origin State '$($oStateRaw)'"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($dCity)) { $invalid++; Write-Host "Skip SAIA DataRow ${currentDataRowForMessage}: Bad Dest City '$($dCityRaw)'"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($dState)) { $invalid++; Write-Host "Skip SAIA DataRow ${currentDataRowForMessage}: Bad Dest State '$($dStateRaw)'"; $skipRow = $true }
            # Numeric Validation
            if (-not $skipRow) { try { $wtNum = [decimal]$wtStr; if($wtNum -le 0){throw} } catch { $invalid++; Write-Host "Skip SAIA DataRow ${currentDataRowForMessage}: Bad Weight '$($wtStrRaw)'"; $skipRow = $true } }
            if (-not $skipRow) { try { $pcsNum = [int]$pcsStr; if($pcsNum -le 0){throw} } catch { $invalid++; Write-Host "Skip SAIA DataRow ${currentDataRowForMessage}: Bad Pieces '$($pcsStrRaw)'"; $skipRow = $true } }
            if (-not $skipRow) { try { $lenNum = [decimal]$lenStr; if($lenNum -le 0){throw} } catch { $invalid++; Write-Host "Skip SAIA DataRow ${currentDataRowForMessage}: Bad Length '$($lenStrRaw)'"; $skipRow = $true } }
            if (-not $skipRow) { try { $widNum = [decimal]$widStr; if($widNum -le 0){throw} } catch { $invalid++; Write-Host "Skip SAIA DataRow ${currentDataRowForMessage}: Bad Width '$($widStrRaw)'"; $skipRow = $true } }
            if (-not $skipRow) { try { $hgtNum = [decimal]$hgtStr; if($hgtNum -le 0){throw} } catch { $invalid++; Write-Host "Skip SAIA DataRow ${currentDataRowForMessage}: Bad Height '$($hgtStrRaw)'"; $skipRow = $true } }

            if ($skipRow) { $rowNum++; continue } # Increment $rowNum here before continuing

            # Create a single commodity item for reports
            # Use the exact property name 'classification' as expected by Invoke-SAIAApi
            # CHANGED: Explicitly cast to [PSCustomObject]
            $commodityItem = [PSCustomObject]@{
                classification = $clStr # Store the trimmed class string
                weight = $wtNum # Store numeric weight
                length = $lenNum
                width = $widNum
                height = $hgtNum
                pieces = $pcsNum
                packagingType = if ($header -contains 'Packaging Type 1') { $row.'Packaging Type 1'.Trim() } else { "PAT" }
                description = if ($header -contains 'Description 1') { $row.'Description 1'.Trim() } else { "Commodity" }
                stackable = if ($header -contains 'Stackable 1') { $row.'Stackable 1'.Trim() } else { "Y" }
            }

            # Create the final normalized object for this row
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
                'Freight Class 1' = $clStr # Store trimmed class string here too
            }
            $normData.Add($newRow)
            $rowNum++ # Increment $rowNum after successful processing of a row
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
        [Parameter(Mandatory=$true)] [array]$Details # Array of hashtables/PSObjects for commodities
    )

    $saiaUserID = $null; $saiaPassword = $null; $saiaRQKey = $null; $accountCodeToUse = $null
    $tariffNameForLog = if ($KeyData.ContainsKey('TariffFileName')) { $KeyData.TariffFileName } elseif ($KeyData.ContainsKey('Name')) {$KeyData.Name} else { "UnknownSAIATariff" }

    # Extract credentials from KeyData
    try {
        if ($KeyData.ContainsKey('UserID')) { $saiaUserID = $KeyData.UserID } else { Write-Host "UserID missing from KeyData for Acct '$tariffNameForLog'."}
        if ($KeyData.ContainsKey('Password')) { $saiaPassword = $KeyData.Password } else { Write-Host "Password missing from KeyData for Acct '$tariffNameForLog'."}
        if ($KeyData.ContainsKey('RQKey')) { $saiaRQKey = $KeyData.RQKey } else { Write-Host "RQKey missing from KeyData for Acct '$tariffNameForLog'."}
        if ($KeyData.ContainsKey('AccountCode')) { $accountCodeToUse = $KeyData.AccountCode } else { Write-Host "AccountCode not found in KeyData for SAIA Acct '$tariffNameForLog'." }
    } catch { Write-Warning "SAIA Credential Extract Fail for Acct '$tariffNameForLog': $($_.Exception.Message)"; return $null }

    # --- Validate Base Inputs ---
    $missingFields = @()
    if ([string]::IsNullOrWhiteSpace($OriginZip)) { $missingFields += "OriginZip" }
    if ([string]::IsNullOrWhiteSpace($DestinationZip)) { $missingFields += "DestinationZip" }
    if ([string]::IsNullOrWhiteSpace($OriginCity)) { $missingFields += "OriginCity" }
    if ([string]::IsNullOrWhiteSpace($OriginState)) { $missingFields += "OriginState" }
    if ([string]::IsNullOrWhiteSpace($DestinationCity)) { $missingFields += "DestinationCity" }
    if ([string]::IsNullOrWhiteSpace($DestinationState)) { $missingFields += "DestinationState" }
    if ($null -eq $Details -or $Details.Count -eq 0) { $missingFields += "Details (Commodity array empty or null)"}
    if ([string]::IsNullOrWhiteSpace($saiaRQKey) -and ([string]::IsNullOrWhiteSpace($saiaUserID) -or [string]::IsNullOrWhiteSpace($saiaPassword))) { $missingFields += "Credentials (Missing RQKey AND UserID/Password pair)" }

    # --- Process Commodity Details Array (Simplified Class Access) ---
    $detailsArrayForPayload = @()
    $totalWeightForPayload = 0.0
    $totalCubeInches = 0.0
    if ($Details -is [array] -and $Details.Count -gt 0) {
        for ($i = 0; $i -lt $Details.Count; $i++) {
            $item = $Details[$i]
            $itemWeight = $null; $itemClass = $null; $itemLength = $null; $itemWidth = $null; $itemHeight = $null; $itemPieces = $null; $itemStackable = 'N'
            $isValidItem = $true
            $currentItemErrors = @() # Track errors for this specific item

            # Check if $item is a usable object type
            if ($item -is [System.Collections.IDictionary] -or $item -is [psobject]) {
                # Try direct property access with validation
                try {
                    # Weight
                    if ($item.PSObject.Properties['weight'] -ne $null -and -not([string]::IsNullOrWhiteSpace($item.weight)) -and ($item.weight -as [decimal]) -ne $null -and [decimal]$item.weight -gt 0) {
                        $itemWeight = [decimal]$item.weight
                    } else { $isValidItem = $false; $currentItemErrors += "Invalid Weight ('$($item.weight)')" }

                    # Class (Simplified Access)
                    $classValue = $null
                    # Check if the 'classification' property exists and try to get its value
                    if ($item -ne $null -and $item.PSObject.Properties.Match('classification').Count -gt 0) { # Using Match as it was in the last provided version
                        $classValue = $item.classification # Access directly
                        Write-Host "DEBUG (Invoke-SAIAApi): Item $($i+1) - Read 'classification' value = [$classValue]"
                    } else {
                         Write-Host "DEBUG (Invoke-SAIAApi): Item $($i+1) - Property 'classification' not found or item is null (using Match)." # Indicate Match was used
                         # If 'classification' is missing, the class is invalid for SAIA
                         $isValidItem = $false
                         $currentItemErrors += "Missing 'classification' property"
                    }

                    # Validate class looks like a number AND is not empty/whitespace (only if property was found)
                    if ($isValidItem) { # Only validate if we think we have a class property
                        if (-not([string]::IsNullOrWhiteSpace($classValue)) -and $classValue -match '^\d+(\.\d+)?$') {
                           $itemClass = $classValue # Keep as string, API payload needs double
                        } else {
                            $isValidItem = $false
                            # Log the value that failed validation, even if it was $null
                            $currentItemErrors += "Invalid Class value ('$($classValue)')"
                        }
                    }

                    # Length
                    if ($item.PSObject.Properties['length'] -ne $null -and -not([string]::IsNullOrWhiteSpace($item.length)) -and ($item.length -as [decimal]) -ne $null -and [decimal]$item.length -gt 0) {
                        $itemLength = [decimal]$item.length
                    } else { $isValidItem = $false; $currentItemErrors += "Invalid Length ('$($item.length)')" }

                    # Width
                    if ($item.PSObject.Properties['width'] -ne $null -and -not([string]::IsNullOrWhiteSpace($item.width)) -and ($item.width -as [decimal]) -ne $null -and [decimal]$item.width -gt 0) {
                        $itemWidth = [decimal]$item.width
                    } else { $isValidItem = $false; $currentItemErrors += "Invalid Width ('$($item.width)')" }

                    # Height
                    if ($item.PSObject.Properties['height'] -ne $null -and -not([string]::IsNullOrWhiteSpace($item.height)) -and ($item.height -as [decimal]) -ne $null -and [decimal]$item.height -gt 0) {
                        $itemHeight = [decimal]$item.height
                    } else { $isValidItem = $false; $currentItemErrors += "Invalid Height ('$($item.height)')" }

                    # Pieces
                    if ($item.PSObject.Properties['pieces'] -ne $null -and -not([string]::IsNullOrWhiteSpace($item.pieces)) -and ($item.pieces -as [int]) -ne $null -and [int]$item.pieces -gt 0) {
                        $itemPieces = [int]$item.pieces
                    } else { $isValidItem = $false; $currentItemErrors += "Invalid Pieces ('$($item.pieces)')" }

                    # Stackable (Optional, defaults to N)
                    if ($item.PSObject.Properties['stackable'] -ne $null -and $item.stackable -eq 'Y') {
                        $itemStackable = 'Y'
                    } # Otherwise remains 'N'

                } catch {
                    # Catch unexpected errors during property access/conversion
                    $isValidItem = $false
                    $currentItemErrors += "Unexpected error accessing properties: $($_.Exception.Message)"
                }

                # If item passed validation, add to payload array
                if ($isValidItem) {
                    # Convert to types expected by SAIA API payload (based on previous code)
                    $detailsArrayForPayload += [ordered]@{
                        length = [double]$itemLength
                        width  = [double]$itemWidth
                        height = [double]$itemHeight
                        weight = [int]$itemWeight # API example showed integer weight
                        class  = [double]$itemClass # API example showed number class
                        units  = $itemPieces
                    }
                    $totalWeightForPayload += $itemWeight
                    $totalCubeInches += ($itemLength * $itemWidth * $itemHeight * $itemPieces)
                } else {
                    # Add collected errors for this item to the main missing fields list
                    $missingFields += "Item $($i+1): $($currentItemErrors -join ', ')"
                }

            } else {
                 # $item was not a dictionary or psobject
                 $isValidItem = $false
                 $missingFields += "Item $($i+1) is not a valid object type (Type: $($item.GetType().FullName))." # Add type info to error
            }
        } # End for loop

        # Check if any valid items were found after iterating through all provided details
        if ($detailsArrayForPayload.Count -eq 0 -and $Details.Count -gt 0) {
             $missingFields += "No valid commodity items found after validation."
        }
    }
    # End Commodity Processing (Simplified Class Access)


    # Check if any errors were added during base input or commodity validation
    if ($missingFields.Count -gt 0) {
        $contextZips = "OZip:$OriginZip DZip:$DestinationZip"
        Write-Warning "SAIA Skip: Acct '$tariffNameForLog' $contextZips - Missing/Invalid required data: $($missingFields -join '; ')." # Use semicolon for readability
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
        }
        destination = @{
            city = $DestinationCity
            state = $DestinationState
            zipcode = $DestinationZip
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
    Write-Host "SAIA Payload ($tariffNameForLog): $payload"

    # --- API Call ---
    $headers = @{ 'Content-Type' = 'application/json'; 'Cache-Control' = 'no-cache' }
    if (-not [string]::IsNullOrWhiteSpace($saiaRQKey)) {
        $headers.'RQ-Key' = $saiaRQKey
    }

    try {
        $apiUrl = $script:saiaApiUri
        if ([string]::IsNullOrWhiteSpace($apiUrl)) { throw "SAIA API URI ($($apiUrl)) is not defined or empty in TMS_Config.ps1."}

        # Final credential check before call
        if (-not $headers.ContainsKey('RQ-Key') -and ([string]::IsNullOrWhiteSpace($payloadObject.userID) -or [string]::IsNullOrWhiteSpace($payloadObject.password)) ) {
             throw "Cannot call SAIA API for Acct '$tariffNameForLog': No RQ-Key header provided AND UserID/Password pair is incomplete/missing in payload."
        }

        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $payload -ErrorAction Stop
        Write-Host "SAIA OK: Acct '$tariffNameForLog'"

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
             # Write-Host "SAIA Full Response: $($response | ConvertTo-Json -Depth 5)" # Uncomment for deep debug
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

Write-Host "TMS SAIA Helper Functions loaded."
