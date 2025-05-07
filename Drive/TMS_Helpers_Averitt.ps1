# TMS_Helpers_Averitt.ps1
# Description: Contains helper functions specific to Averitt Express operations,
#              including data normalization and API interaction based on the Dynamic Pricing API guide.
#              This file should be dot-sourced by the main script(s) after TMS_Config.ps1.

# Assumes config variables like $script:averittApiUri are available from TMS_Config.ps1
# Assumes general helper functions (if any were used by these) are available.

# --- Data Normalization Functions ---

function Load-And-Normalize-AverittData {
    param(
        [Parameter(Mandatory=$true)][string]$CsvPath
    )
    Write-Host "`nLoading Averitt data: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan

    # Define required and optional columns based on Averitt API needs (Dynamic Pricing)
    # Required for basic structure: Origin/Dest City, State, Zip, Country, BillTo Name/Address/City/State/Zip/Country
    # Required for commodities: Weight, Length, Width, Height, Pieces, PackagingType (PAT/LSE), Classification (optional but needed for LTL)
    # Optional: Origin/Dest Account, PickupDate, Accessorials, Hazmat contact, Commodity Description, Stackable
    $reqBaseCols = @(
        "Origin City", "Origin State", "Origin Postal Code", "Origin Country",
        "Destination City", "Destination State", "Destination Postal Code", "Destination Country"
        # BillTo info is often static or derived, but check if needed from CSV
        # "BillTo Name", "BillTo Address", "BillTo City", "BillTo State", "BillTo Zip", "BillTo Country"
    )
    # Commodity columns - Assuming up to 5 commodities for flexibility, checking for at least 1
    $reqCommColsBase = @("Freight Class 1", "Total Weight", "Total Length", "Total Width", "Total Height", "Total Units") # Map CSV totals to Commodity 1
    $optCommColsBase = @("Packaging Type 1", "Description 1", "Stackable 1") # Add more if CSV has them per commodity

    # Check if essential base columns exist
    try {
        if (-not (Test-Path -Path $CsvPath -PathType Leaf)) {
            Write-Error "CSV file not found at '$CsvPath'."
            return $null
        }
        # Read just the header first
        $header = (Import-Csv -Path $CsvPath -TotalCount 1).PSObject.Properties.Name
        if ($null -eq $header) { Write-Error "Could not read header from CSV '$CsvPath'."; return $null }

        $missingReqBase = $reqBaseCols | Where-Object { $_ -notin $header }
        if ($missingReqBase.Count -gt 0) { Write-Error "CSV missing required Averitt base columns: $($missingReqBase -join ', ')"; return $null }

        # Check for at least the first set of required commodity columns
        $missingReqComm = $reqCommColsBase | Where-Object { $_ -notin $header }
        if ($missingReqComm.Count -gt 0) { Write-Error "CSV missing required Averitt commodity 1 columns: $($missingReqComm -join ', ')"; return $null }

        # Now import the full data
        $rawData = Import-Csv -Path $CsvPath -ErrorAction Stop
        Write-Host " -> Rows read from CSV: $($rawData.Count)." -ForegroundColor Gray
        if ($rawData.Count -eq 0) { Write-Warning "CSV empty."; return @() }

        $normData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing Averitt data..." -ForegroundColor Gray
        $invalidRowCount = 0; $rowNum = 1
        foreach ($row in $rawData) {
            $rowNum++
            $normalizedEntry = [ordered]@{ # Using ordered hashtable for clarity
                # Service (Defaults based on API guide)
                ServiceLevel = "STND"
                # Payment (Defaults - Assume Shipper Prepaid unless CSV indicates otherwise)
                PaymentTerms = "Prepaid"
                PaymentPayer = "Shipper" # This determines where the Account# is required
                # Transit (Default to today if not in CSV)
                PickupDate = (Get-Date -Format 'yyyyMMdd')
                # Origin
                OriginAccount = $null # Required if Payer=Shipper, needs lookup or CSV column
                OriginCity = $row."Origin City".Trim()
                OriginStateProvince = $row."Origin State".Trim()
                OriginPostalCode = $row."Origin Postal Code".Trim()
                OriginCountry = $row."Origin Country".Trim()
                # Destination
                DestinationAccount = $null # Required if Payer=Consignee
                DestinationCity = $row."Destination City".Trim()
                DestinationStateProvince = $row."Destination State".Trim()
                DestinationPostalCode = $row."Destination Postal Code".Trim()
                DestinationCountry = $row."Destination Country".Trim()
                # BillTo (Defaults - Required by API, may need real data)
                BillToAccount = $null # Required if Payer=ThirdParty
                BillToName = "Default BillTo" # Placeholder - Get from CSV if available
                BillToAddress = "123 Default St" # Placeholder
                BillToCity = "DefaultCity" # Placeholder
                BillToStateProvince = "TN" # Placeholder
                BillToPostalCode = "00000" # Placeholder
                BillToCountry = "USA" # Placeholder
                # Commodities Array
                Commodities = @()
                # Accessorials (Optional)
                Accessorials = $null # Will build if needed
            }

            # --- Populate Commodity Details ---
            $commodity = [ordered]@{
                classification = $row."Freight Class 1".Trim()
                weight = $row."Total Weight".Trim() # API expects string
                length = $row."Total Length".Trim() # API expects string
                width = $row."Total Width".Trim() # API expects string
                height = $row."Total Height".Trim() # API expects string
                pieces = $row."Total Units".Trim() # API expects string
                packagingType = "PAT" # Default to Pallet, check CSV if available
                description = "Commodity" # Default, check CSV
                stackable = "Y" # Default, check CSV
            }
            # Add optional commodity fields if present in CSV header
            if ($header -contains "Packaging Type 1") { $commodity.packagingType = $row."Packaging Type 1".Trim() } # Adjust column name as needed
            if ($header -contains "Description 1") { $commodity.description = $row."Description 1".Trim() }
            if ($header -contains "Stackable 1") { $commodity.stackable = $row."Stackable 1".Trim() } # Expects Y/N

            # Basic validation for commodity 1
            $commValid = $true
            if ([string]::IsNullOrWhiteSpace($commodity.weight) -or ([decimal]$commodity.weight -le 0)) { $commValid = $false; Write-Verbose "Skip Row $rowNum: Invalid weight '$($commodity.weight)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.length) -or ([decimal]$commodity.length -le 0)) { $commValid = $false; Write-Verbose "Skip Row $rowNum: Invalid length '$($commodity.length)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.width) -or ([decimal]$commodity.width -le 0)) { $commValid = $false; Write-Verbose "Skip Row $rowNum: Invalid width '$($commodity.width)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.height) -or ([decimal]$commodity.height -le 0)) { $commValid = $false; Write-Verbose "Skip Row $rowNum: Invalid height '$($commodity.height)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.pieces) -or ([int]$commodity.pieces -le 0)) { $commValid = $false; Write-Verbose "Skip Row $rowNum: Invalid pieces '$($commodity.pieces)'" }
            # Classification is optional in API but usually needed for LTL rating
            if ([string]::IsNullOrWhiteSpace($commodity.classification)) { Write-Verbose "Row $rowNum: Missing classification for commodity 1." }

            if ($commValid) {
                $normalizedEntry.Commodities += $commodity
            } else {
                $invalidRowCount++; continue # Skip row if essential commodity data is bad
            }

            # --- Add Logic for More Commodities if CSV supports it ---
            # Example for Commodity 2:
            # if ($header -contains "Freight Class 2" -and -not [string]::IsNullOrWhiteSpace($row."Freight Class 2")) {
            #    $commodity2 = [ordered]@{ classification=...; weight=...; ... }
            #    # Validate commodity2...
            #    if ($comm2Valid) { $normalizedEntry.Commodities += $commodity2 }
            # }

            # --- Add Logic for Accessorials if CSV supports it ---
            # Example:
            # if ($header -contains "Liftgate Delivery" -and $row."Liftgate Delivery" -match '^(true|Y|1)$') {
            #    if ($null -eq $normalizedEntry.Accessorials) { $normalizedEntry.Accessorials = @{ codes = @() } }
            #    $normalizedEntry.Accessorials.codes += "LFTD"
            # }
            # if ($header -contains "Hazmat" -and $row."Hazmat" -match '^(true|Y|1)$') {
            #    if ($null -eq $normalizedEntry.Accessorials) { $normalizedEntry.Accessorials = @{ codes = @() } }
            #    $normalizedEntry.Accessorials.codes += "HAZ"
            #    # Add hazmat contact info if available
            #    if ($header -contains "Hazmat Contact" -and $header -contains "Hazmat Phone") {
            #       $normalizedEntry.Accessorials.hazardousContact = @{
            #           name = $row."Hazmat Contact"
            #           phone = $row."Hazmat Phone" -replace '[^0-9]' # Clean phone
            #       }
            #    }
            # }

            # --- Basic Validation for Base Fields ---
            $baseValid = $true
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.OriginCity)) { $baseValid = $false }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.OriginStateProvince)) { $baseValid = $false }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.OriginPostalCode)) { $baseValid = $false }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.OriginCountry)) { $baseValid = $false }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.DestinationCity)) { $baseValid = $false }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.DestinationStateProvince)) { $baseValid = $false }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.DestinationPostalCode)) { $baseValid = $false }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.DestinationCountry)) { $baseValid = $false }
            # Add validation for BillTo fields if they are expected from CSV

            if (-not $baseValid) {
                Write-Warning "Skipping row $rowNum due to missing required Origin/Destination fields."
                $invalidRowCount++; continue
            }

            # Add the fully populated and validated entry
            $normData.Add([PSCustomObject]$normalizedEntry)

        } # End foreach row

        if ($invalidRowCount -gt 0) { Write-Warning " -> Skipped $invalidRowCount Averitt rows (missing/invalid essential data)." }
        Write-Host " -> OK: $($normData.Count) Averitt rows normalized." -ForegroundColor Green
        return $normData
    } catch {
        Write-Error "Error processing Averitt CSV '$CsvPath': $($_.Exception.Message)"; return $null
    }
}


# --- API Call Functions ---

function Invoke-AverittApi {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$KeyData, # Contains APIKey, potentially default Account# if needed

        [Parameter(Mandatory=$true)]
        [PSCustomObject]$ShipmentData # Normalized data from Load-And-Normalize-AverittData
    )

    $apiKey = $KeyData.APIKey
    $tariffNameForLog = if ($KeyData.ContainsKey('TariffFileName')) { $KeyData.TariffFileName } else { "UnknownAverittTariff" }

    # --- Input Validation ---
    if ([string]::IsNullOrWhiteSpace($apiKey)) { Write-Warning "Averitt Skip ($tariffNameForLog): APIKey missing from KeyData."; return $null }
    if ($null -eq $ShipmentData) { Write-Warning "Averitt Skip ($tariffNameForLog): ShipmentData object is null."; return $null }
    if ($null -eq $ShipmentData.Commodities -or $ShipmentData.Commodities.Count -eq 0) { Write-Warning "Averitt Skip ($tariffNameForLog): No valid commodities found in ShipmentData."; return $null }
    # Add more validation for required fields in $ShipmentData (Origin/Dest Zip, City, State, Country etc.)

    # --- Construct Payload ---
    # Map normalized $ShipmentData properties to the API JSON structure
    $payloadObject = [ordered]@{
        service = @{
            level = $ShipmentData.ServiceLevel # Should be "STND"
        }
        payment = @{
            terms = $ShipmentData.PaymentTerms   # Prepaid, Collect, Third Party
            payer = $ShipmentData.PaymentPayer   # Shipper, Consignee, Third Party
        }
        transit = @{
            pickupDate = $ShipmentData.PickupDate # YYYYMMDD
        }
        commodities = @() # Initialize as array
        origin = @{
            account = $ShipmentData.OriginAccount # Required if payer=Shipper
            city = $ShipmentData.OriginCity
            stateProvince = $ShipmentData.OriginStateProvince
            postalCode = $ShipmentData.OriginPostalCode
            country = $ShipmentData.OriginCountry # USA, CAN, MEX
        }
        destination = @{
            account = $ShipmentData.DestinationAccount # Required if payer=Consignee
            city = $ShipmentData.DestinationCity
            stateProvince = $ShipmentData.DestinationStateProvince
            postalCode = $ShipmentData.DestinationPostalCode
            country = $ShipmentData.DestinationCountry # USA, CAN, MEX
        }
        billTo = @{ # Required by API
            account = $ShipmentData.BillToAccount # Required if payer=ThirdParty
            name = $ShipmentData.BillToName
            address = $ShipmentData.BillToAddress
            city = $ShipmentData.BillToCity
            stateProvince = $ShipmentData.BillToStateProvince
            postalCode = $ShipmentData.BillToPostalCode
            country = $ShipmentData.BillToCountry
        }
    }

    # Add commodities from the normalized data
    foreach ($comm in $ShipmentData.Commodities) {
        $payloadObject.commodities += $comm # Assumes $comm is already an ordered hashtable/PSObject
    }

    # Add accessorials if present in normalized data
    if ($null -ne $ShipmentData.Accessorials) {
        $payloadObject.accessorials = $ShipmentData.Accessorials
        # Ensure nulls are represented correctly if sub-objects are not present
        if (-not $payloadObject.accessorials.ContainsKey('hazardousContact')) { $payloadObject.accessorials.hazardousContact = $null }
        if (-not $payloadObject.accessorials.ContainsKey('cod')) { $payloadObject.accessorials.cod = $null }
        if (-not $payloadObject.accessorials.ContainsKey('insuranceDetails')) { $payloadObject.accessorials.insuranceDetails = $null }
        if (-not $payloadObject.accessorials.ContainsKey('sortAndSegregateDetails')) { $payloadObject.accessorials.sortAndSegregateDetails = $null }
        if (-not $payloadObject.accessorials.ContainsKey('markDetails')) { $payloadObject.accessorials.markDetails = $null }
    } else {
         # API might require the accessorials object even if empty, or with codes=[]. Check API behavior.
         # For safety, let's add an empty one if none were provided.
         $payloadObject.accessorials = @{ codes = @(); hazardousContact=$null; cod=$null; insuranceDetails=$null; sortAndSegregateDetails=$null; markDetails=$null }
    }


    # --- Convert to JSON ---
    $payloadJson = $payloadObject | ConvertTo-Json -Depth 10
    Write-Verbose "Averitt Payload ($tariffNameForLog): $payloadJson"

    # --- API Call ---
    $apiUrl = $script:averittApiUri # Get from config
    if ([string]::IsNullOrWhiteSpace($apiUrl)) { Write-Error "Averitt API URI not configured in TMS_Config.ps1."; return $null }

    $uri = "{0}?api_key={1}" -f $apiUrl, $apiKey
    $headers = @{
        'Content-Type' = 'application/json'
        'Accept'       = 'application/json'
    }

    Write-Verbose "Calling Averitt API: $tariffNameForLog"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $payloadJson -ErrorAction Stop

        # --- Process Response ---
        Write-Verbose "Averitt Raw Response ($tariffNameForLog): $($response | ConvertTo-Json -Depth 5)"

        # Check for successful response structure (based on API Guide example)
        if ($response -ne $null -and $response.messageStatus -ne $null) {
            # Check explicit status
            if ($response.messageStatus.status -eq 'PASS' -or $response.messageStatus.status -eq 'WARNING') {
                 # Find the total charge - API guide shows it in shipmentInfo.totalCharge
                 # BUT the example response shows charges broken down in accessorials, and shipmentInfo.totalCharge is 0.00
                 # We need the NET charge. Let's assume the API *should* return a reliable total somewhere.
                 # If shipmentInfo.totalCharge is unreliable, we might need to sum charges manually (less ideal).
                 # Let's TRY shipmentInfo.totalCharge first.

                 $totalChargeValue = $null
                 if ($response.shipmentInfo -ne $null -and $response.shipmentInfo.PSObject.Properties.Name -contains 'totalCharge') {
                      $totalChargeValue = $response.shipmentInfo.totalCharge
                 }

                 # Fallback: Sum accessorial charges if totalCharge is zero or missing (This is a guess based on example)
                 if (($null -eq $totalChargeValue -or [decimal]$totalChargeValue -eq 0) -and $response.accessorials -ne $null -and $response.accessorials -is [array]) {
                     Write-Warning "Averitt ($tariffNameForLog): shipmentInfo.totalCharge was zero or null. Attempting to sum accessorial charges as a fallback."
                     $calculatedTotal = 0.0
                     foreach ($acc in $response.accessorials) {
                         if ($acc.chargeAmount -ne $null -and $acc.chargeAmount -as [decimal] -ne $null) {
                             $calculatedTotal += [decimal]$acc.chargeAmount
                         }
                     }
                     # Need base freight charge too - where is it in the response? API guide is unclear.
                     # This fallback is likely incomplete without knowing where base charge is returned.
                     # For now, we'll stick to trying totalCharge.
                     Write-Warning "Averitt ($tariffNameForLog): Fallback sum is likely incomplete as base freight charge location in response is unclear."
                 }


                 if ($totalChargeValue -ne $null) {
                     try {
                         $cleanedRate = $totalChargeValue -replace '[$,]' # Remove currency symbols if any
                         $decimalRate = [decimal]$cleanedRate
                         if ($decimalRate -lt 0) { throw "Rate cannot be negative."} # Basic sanity check
                         Write-Verbose "Averitt OK ($tariffNameForLog): Rate = $decimalRate"
                         return $decimalRate
                     } catch {
                         Write-Warning "Averitt Convert Fail ($tariffNameForLog): Cannot convert rate '$totalChargeValue' to valid decimal. Error: $($_.Exception.Message)"; return $null
                     }
                 } else {
                     Write-Warning "Averitt Rate Not Found ($tariffNameForLog): Could not find 'shipmentInfo.totalCharge' in the successful/warning response."
                     return $null
                 }

            } else { # Status is FAIL
                 Write-Warning "Averitt API Fail ($tariffNameForLog): Status=$($response.messageStatus.status), Code=$($response.messageStatus.code), Msg=$($response.messageStatus.message)"
                 return $null
            }
        } else {
             Write-Warning "Averitt API Error ($tariffNameForLog): Response structure invalid or missing 'messageStatus'."
             return $null
        }

    } catch {
        $errMsg = $_.Exception.Message; $statusCode = "N/A"; $eBody = "N/A"
        if ($_.Exception.Response) {
             try {$statusCode = $_.Exception.Response.StatusCode.value__} catch{}
             try { $stream = $_.Exception.Response.GetResponseStream(); $reader = New-Object System.IO.StreamReader($stream); $eBody = $reader.ReadToEnd(); $reader.Close(); $stream.Close() } catch {$eBody="(Err reading resp body: $($_.Exception.Message))"}
        }
        $truncatedBody = if ($eBody.Length -gt 500) { $eBody.Substring(0, 500) + "..." } else { $eBody }
        $fullErrMsg = "Averitt Invoke FAIL ($tariffNameForLog): Error: $errMsg (HTTP $statusCode) Resp: $truncatedBody"
        Write-Warning $fullErrMsg;
        return $null
    }
}

Write-Verbose "TMS Averitt Helper Functions loaded."