# TMS_Helpers_Averitt.ps1
# Description: Contains helper functions specific to Averitt Express operations,
#              including data normalization and API interaction based on the Dynamic Pricing API guide.
#              Compatible with PowerShell 5.1+.
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
    $reqBaseCols = @(
        "Origin City", "Origin State", "Origin Postal Code", "Origin Country",
        "Destination City", "Destination State", "Destination Postal Code", "Destination Country"
    )
    # Commodity columns - Map CSV totals to Commodity 1 for reports
    $reqCommColsBase = @("Freight Class 1", "Total Weight", "Total Length", "Total Width", "Total Height", "Total Units")
    $optCommColsBase = @("Packaging Type 1", "Description 1", "Stackable 1")

    try {
        if (-not (Test-Path -Path $CsvPath -PathType Leaf)) { Write-Error "CSV file not found: '$CsvPath'."; return $null }

        # MODIFIED SECTION: Use Import-Csv for the whole file, then get header
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
        # END MODIFIED SECTION

        Write-Host " -> Rows read from CSV: $($rawData.Count)." -ForegroundColor Gray

        # Check for required columns using the extracted header
        $missingReqBase = $reqBaseCols | Where-Object { $_ -notin $header }
        if ($missingReqBase.Count -gt 0) { Write-Error "CSV missing Averitt base columns: $($missingReqBase -join ', ')"; return $null }
        $missingReqComm = $reqCommColsBase | Where-Object { $_ -notin $header }
        if ($missingReqComm.Count -gt 0) { Write-Error "CSV missing Averitt commodity 1 columns: $($missingReqComm -join ', ')"; return $null }

        $missingOpt = $optCommColsBase | Where-Object { $_ -notin $header }
        if($missingOpt.Count -gt 0){ Write-Warning "CSV missing optional Averitt columns: $($missingOpt -join ', ')" }

        # --- Proceed with Normalization using $rawData ---
        $normData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing Averitt data..." -ForegroundColor Gray
        $invalidRowCount = 0; $rowNum = 1 # Start rowNum at 1 for messages (header is row 0 effectively)
        foreach ($row in $rawData) {
            $rowNum++ # Increment for current data row number

            $normalizedEntry = [ordered]@{
                ServiceLevel = "STND"; PaymentTerms = "Prepaid"; PaymentPayer = "Shipper"; # Defaults
                PickupDate = (Get-Date -Format 'yyyyMMdd'); # Default
                OriginAccount = 1418543
                OriginCity = $row."Origin City".Trim(); OriginStateProvince = $row."Origin State".Trim(); OriginPostalCode = $row."Origin Postal Code".Trim(); OriginCountry = $row."Origin Country".Trim();
                DestinationAccount = $null
                DestinationCity = $row."Destination City".Trim(); DestinationStateProvince = $row."Destination State".Trim(); DestinationPostalCode = $row."Destination Postal Code".Trim(); DestinationCountry = $row."Destination Country".Trim();
                BillToAccount = $null
                BillToName = "Default BillTo"; BillToAddress = "123 Default St"; BillToCity = "DefaultCity"; BillToStateProvince = "TN"; BillToPostalCode = "00000"; BillToCountry = "USA"; # Defaults
                Commodities = @(); Accessorials = $null
            }

            # --- Populate Commodity Details ---
            $commodity = [ordered]@{
                classification = $row."Freight Class 1".Trim()
                weight = $row."Total Weight".Trim()
                length = $row."Total Length".Trim()
                width = $row."Total Width".Trim()
                height = $row."Total Height".Trim()
                pieces = $row."Total Units".Trim()
                packagingType = if ($header -contains "Packaging Type 1") { $row."Packaging Type 1".Trim() } else { "PAT" }
                description = if ($header -contains "Description 1") { $row."Description 1".Trim() } else { "Commodity" }
                stackable = if ($header -contains "Stackable 1") { $row."Stackable 1".Trim() } else { "Y" }
            }

            # Basic validation for commodity 1
            $commValid = $true
            if ([string]::IsNullOrWhiteSpace($commodity.weight) -or ($commodity.weight -as [decimal]) -eq $null -or ([decimal]$commodity.weight -le 0)) { $commValid = $false; Write-Host "Skip Row ${rowNum}: Invalid weight '$($commodity.weight)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.length) -or ($commodity.length -as [decimal]) -eq $null -or ([decimal]$commodity.length -le 0)) { $commValid = $false; Write-Host "Skip Row ${rowNum}: Invalid length '$($commodity.length)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.width) -or ($commodity.width -as [decimal]) -eq $null -or ([decimal]$commodity.width -le 0)) { $commValid = $false; Write-Host "Skip Row ${rowNum}: Invalid width '$($commodity.width)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.height) -or ($commodity.height -as [decimal]) -eq $null -or ([decimal]$commodity.height -le 0)) { $commValid = $false; Write-Host "Skip Row ${rowNum}: Invalid height '$($commodity.height)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.pieces) -or ($commodity.pieces -as [int]) -eq $null -or ([int]$commodity.pieces -le 0)) { $commValid = $false; Write-Host "Skip Row ${rowNum}: Invalid pieces '$($commodity.pieces)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.classification)) { Write-Host "Row ${rowNum}: Missing classification for commodity 1." } # Allow if missing, but log

            if ($commValid) {
                $normalizedEntry.Commodities += $commodity
            } else {
                $invalidRowCount++; continue # Skip row if essential commodity data is bad
            }

            # --- Basic Validation for Base Fields ---
            $baseValid = $true
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.OriginCity)) { $baseValid = $false; Write-Host "Skip Row ${rowNum}: Missing Origin City" }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.OriginStateProvince)) { $baseValid = $false; Write-Host "Skip Row ${rowNum}: Missing Origin State" }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.OriginPostalCode)) { $baseValid = $false; Write-Host "Skip Row ${rowNum}: Missing Origin Postal Code" }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.OriginCountry)) { $baseValid = $false; Write-Host "Skip Row ${rowNum}: Missing Origin Country" }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.DestinationCity)) { $baseValid = $false; Write-Host "Skip Row ${rowNum}: Missing Destination City" }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.DestinationStateProvince)) { $baseValid = $false; Write-Host "Skip Row ${rowNum}: Missing Destination State" }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.DestinationPostalCode)) { $baseValid = $false; Write-Host "Skip Row ${rowNum}: Missing Destination Postal Code" }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.DestinationCountry)) { $baseValid = $false; Write-Host "Skip Row ${rowNum}: Missing Destination Country" }


            if (-not $baseValid) {
                Write-Warning "Skipping data row $rowNum due to missing essential Origin/Destination fields after trimming."
                $invalidRowCount++
                continue
            }

            $normData.Add([PSCustomObject]$normalizedEntry)
        } # End foreach row

        if ($invalidRowCount -gt 0) { Write-Warning " -> Skipped $invalidRowCount Averitt rows due to validation errors." }
        Write-Host " -> OK: $($normData.Count) Averitt rows normalized." -ForegroundColor Green
        return $normData
    } catch {
        # Catch errors from Test-Path, Import-Csv, or within the loop
        Write-Error "Error processing Averitt CSV '$CsvPath': $($_.Exception.Message)"
        return $null
    }
}


# --- API Call Functions ---

function Invoke-AverittApi {
    param(
        [Parameter(Mandatory=$true)] [hashtable]$KeyData,
        [Parameter(Mandatory=$true)] [PSCustomObject]$ShipmentData # Normalized data object from Load-And-Normalize-AverittData
    )

    $apiKey = $KeyData.APIKey
    $tariffNameForLog = if ($KeyData.TariffFileName) { $KeyData.TariffFileName } elseif ($KeyData.Name) {$KeyData.Name} else { "UnknownAverittTariff" }

    # --- Input Validation ---
    if ([string]::IsNullOrWhiteSpace($apiKey)) { Write-Warning "Averitt Skip ($tariffNameForLog): APIKey missing from KeyData."; return $null }
    if ($null -eq $ShipmentData) { Write-Warning "Averitt Skip ($tariffNameForLog): ShipmentData object is null."; return $null }
    if ($null -eq $ShipmentData.Commodities -or $ShipmentData.Commodities.Count -eq 0) { Write-Warning "Averitt Skip ($tariffNameForLog): No commodities found in ShipmentData."; return $null }

    # --- Determine Required Account Number based on Payer ---
    $originAccountNum = if ($ShipmentData.PSObject.Properties.Name -contains 'OriginAccount' -and -not [string]::IsNullOrWhiteSpace($ShipmentData.OriginAccount)) { $ShipmentData.OriginAccount } elseif ($KeyData.ContainsKey('AccountNumber')) { $KeyData.AccountNumber } else { $null }
    $destAccountNum = if ($ShipmentData.PSObject.Properties.Name -contains 'DestinationAccount' -and -not [string]::IsNullOrWhiteSpace($ShipmentData.DestinationAccount)) { $ShipmentData.DestinationAccount } elseif ($KeyData.ContainsKey('AccountNumber')) { $KeyData.AccountNumber } else { $null }
    $billToAccountNum = if ($ShipmentData.PSObject.Properties.Name -contains 'BillToAccount' -and -not [string]::IsNullOrWhiteSpace($ShipmentData.BillToAccount)) { $ShipmentData.BillToAccount } elseif ($KeyData.ContainsKey('AccountNumber')) { $KeyData.AccountNumber } else { $null }

    $payer = $ShipmentData.PaymentPayer
    if ($payer -eq 'Shipper' -and [string]::IsNullOrWhiteSpace($originAccountNum)) { Write-Warning "Averitt Skip ($tariffNameForLog): Payer is Shipper but Origin Account number is missing (checked ShipmentData.OriginAccount and KeyData.AccountNumber)."; return $null }
    if ($payer -eq 'Consignee' -and [string]::IsNullOrWhiteSpace($destAccountNum)) { Write-Warning "Averitt Skip ($tariffNameForLog): Payer is Consignee but Destination Account number is missing (checked ShipmentData.DestinationAccount and KeyData.AccountNumber)."; return $null }
    if ($payer -eq 'ThirdParty' -and [string]::IsNullOrWhiteSpace($billToAccountNum)) { Write-Warning "Averitt Skip ($tariffNameForLog): Payer is ThirdParty but BillTo Account number is missing (checked ShipmentData.BillToAccount and KeyData.AccountNumber)."; return $null }


    # --- Construct Payload ---
    $payloadObject = [ordered]@{
        service = @{ level = $ShipmentData.ServiceLevel }
        payment = @{ terms = $ShipmentData.PaymentTerms; payer = $payer }
        transit = @{ pickupDate = $ShipmentData.PickupDate }
        commodities = @()
        origin = @{
            account = $originAccountNum
            city = $ShipmentData.OriginCity
            stateProvince = $ShipmentData.OriginStateProvince
            postalCode = $ShipmentData.OriginPostalCode
            country = $ShipmentData.OriginCountry
        }
        destination = @{
            account = $destAccountNum
            city = $ShipmentData.DestinationCity
            stateProvince = $ShipmentData.DestinationStateProvince
            postalCode = $ShipmentData.DestinationPostalCode
            country = $ShipmentData.DestinationCountry
        }
        billTo = @{
            account = $billToAccountNum
            name = $ShipmentData.BillToName
            address = $ShipmentData.BillToAddress
            city = $ShipmentData.BillToCity
            stateProvince = $ShipmentData.BillToStateProvince
            postalCode = $ShipmentData.BillToPostalCode
            country = $ShipmentData.BillToCountry
        }
    }
    foreach ($commIn in $ShipmentData.Commodities) {
        # Ensure all commodity fields are strings as per API examples
        $apiComm = [ordered]@{
            classification = [string]$commIn.classification
            weight = [string]$commIn.weight
            length = [string]$commIn.length
            width = [string]$commIn.width
            height = [string]$commIn.height
            pieces = [string]$commIn.pieces
            packagingType = [string]$commIn.packagingType
            description = [string]$commIn.description
            stackable = [string]$commIn.stackable # Should be 'Y' or 'N'
        }
        $payloadObject.commodities += $apiComm
    }

    # Ensure accessorials structure is present even if empty, with all optional sub-objects for safety
    if ($null -ne $ShipmentData.Accessorials -and $ShipmentData.Accessorials.PSObject.Properties.Name -contains 'codes') {
        $payloadObject.accessorials = $ShipmentData.Accessorials
    } else {
        $payloadObject.accessorials = @{ codes = @() } # Default to empty codes array
    }
    # Ensure optional complex objects within accessorials are present if accessorials itself is defined
    if ($payloadObject.PSObject.Properties.Name -contains 'accessorials' -and $null -ne $payloadObject.accessorials) {
        if (-not $payloadObject.accessorials.PSObject.Properties.Name -contains 'hazardousContact') { $payloadObject.accessorials.hazardousContact = $null }
        if (-not $payloadObject.accessorials.PSObject.Properties.Name -contains 'cod') { $payloadObject.accessorials.cod = $null }
        if (-not $payloadObject.accessorials.PSObject.Properties.Name -contains 'insuranceDetails') { $payloadObject.accessorials.insuranceDetails = $null }
        if (-not $payloadObject.accessorials.PSObject.Properties.Name -contains 'sortAndSegregateDetails') { $payloadObject.accessorials.sortAndSegregateDetails = $null }
        if (-not $payloadObject.accessorials.PSObject.Properties.Name -contains 'markDetails') { $payloadObject.accessorials.markDetails = $null }
    }


    # --- Convert to JSON ---
    $payloadJson = $payloadObject | ConvertTo-Json -Depth 10
    Write-Host "Averitt Payload ($tariffNameForLog): $payloadJson"

    # --- API Call ---
    $apiUrl = $script:averittApiUri
    if ([string]::IsNullOrWhiteSpace($apiUrl)) { Write-Error "Averitt API URI not configured in TMS_Config.ps1."; return $null }
    $uri = "{0}?api_key={1}" -f $apiUrl, $apiKey
    $headers = @{ 'Content-Type' = 'application/json'; 'Accept' = 'application/json' }

    Write-Host "Calling Averitt API for tariff: $tariffNameForLog"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $payloadJson -ErrorAction Stop
        Write-Host "Averitt Raw Response ($tariffNameForLog): $($response | ConvertTo-Json -Depth 5)"

        # --- Process Response ---
        if ($response -ne $null -and $response.PSObject.Properties.Name -contains 'messageStatus' -and $response.messageStatus -ne $null) {
            if ($response.messageStatus.status -in @('PASS', 'WARNING')) { # 'WARNING' might still contain a rate
                 $totalChargeValue = $null
                 if ($response.PSObject.Properties.Name -contains 'shipmentInfo' -and $response.shipmentInfo -ne $null -and $response.shipmentInfo.PSObject.Properties.Name -contains 'totalCharge') {
                      $totalChargeValue = $response.shipmentInfo.totalCharge
                 }

                 if ($totalChargeValue -ne $null) {
                     try {
                         $cleanedRate = $totalChargeValue -replace '[$,]'; $decimalRate = [decimal]$cleanedRate
                         if ($decimalRate -lt 0) { throw "Negative rate received from API."}
                         Write-Host "Averitt API Call OK ($tariffNameForLog): Rate = $decimalRate"
                         return $decimalRate
                     } catch { Write-Warning "Averitt Rate Conversion Fail ($tariffNameForLog): Cannot convert rate '$totalChargeValue' to decimal. Error: $($_.Exception.Message)"; return $null }
                 } else { Write-Warning "Averitt Rate Not Found in Response ($tariffNameForLog): 'shipmentInfo.totalCharge' field is missing or null in API response, even with PASS/WARNING status."; return $null }
            } else { Write-Warning "Averitt API Call Failed ($tariffNameForLog): Status=$($response.messageStatus.status), Code=$($response.messageStatus.code), Message=$($response.messageStatus.message)"; return $null }
        } else { Write-Warning "Averitt API Call Error ($tariffNameForLog): API response was invalid or missing 'messageStatus' field."; return $null }
    } catch {
        $errMsg = $_.Exception.Message; $statusCode = "N/A"; $eBody = "N/A"
        if ($_.Exception.Response) {
            try {$statusCode = $_.Exception.Response.StatusCode.value__} catch{}
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $eBody = $reader.ReadToEnd()
                $reader.Close(); $stream.Close()
            } catch {$eBody="(Error reading response body: $($_.Exception.Message))"}
        }
        $truncatedBody = if ($eBody.Length -gt 500) { $eBody.Substring(0, 500) + "..." } else { $eBody }
        $fullErrMsg = "Averitt API Invoke-RestMethod FAILED ($tariffNameForLog): Error: $errMsg (HTTP Status: $statusCode) Response Body: $truncatedBody"; Write-Warning $fullErrMsg; return $null
    }
}

Write-Host "TMS Averitt Helper Functions loaded."
