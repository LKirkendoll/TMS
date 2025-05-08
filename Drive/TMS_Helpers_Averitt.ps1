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
        $header = (Import-Csv -Path $CsvPath -TotalCount 1).PSObject.Properties.Name
        if ($null -eq $header) { Write-Error "Could not read header from CSV '$CsvPath'."; return $null }

        $missingReqBase = $reqBaseCols | Where-Object { $_ -notin $header }
        if ($missingReqBase.Count -gt 0) { Write-Error "CSV missing Averitt base columns: $($missingReqBase -join ', ')"; return $null }
        $missingReqComm = $reqCommColsBase | Where-Object { $_ -notin $header }
        if ($missingReqComm.Count -gt 0) { Write-Error "CSV missing Averitt commodity 1 columns: $($missingReqComm -join ', ')"; return $null }
        $missingOpt = $optCommColsBase | Where-Object { $_ -notin $header }
        if($missingOpt.Count -gt 0){ Write-Warning "CSV missing optional Averitt columns: $($missingOpt -join ', ')" }

        $rawData = Import-Csv -Path $CsvPath -ErrorAction Stop
        Write-Host " -> Rows read from CSV: $($rawData.Count)." -ForegroundColor Gray
        if ($rawData.Count -eq 0) { Write-Warning "CSV empty."; return @() }

        $normData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing Averitt data..." -ForegroundColor Gray
        $invalidRowCount = 0; $rowNum = 1
        foreach ($row in $rawData) {
            $rowNum++
            # <<< Consider adding columns for Account numbers if they vary per shipment in CSV >>>
            # Example: $originAccountFromCsv = if ($header -contains 'Origin Account') { $row.'Origin Account'.Trim() } else { $null }
            $normalizedEntry = [ordered]@{
                ServiceLevel = "STND"; PaymentTerms = "Prepaid"; PaymentPayer = "Shipper"; # Defaults
                PickupDate = (Get-Date -Format 'yyyyMMdd'); # Default
                OriginAccount = $null # Placeholder - Needs population based on Payer
                OriginCity = $row."Origin City".Trim(); OriginStateProvince = $row."Origin State".Trim(); OriginPostalCode = $row."Origin Postal Code".Trim(); OriginCountry = $row."Origin Country".Trim();
                DestinationAccount = $null # Placeholder
                DestinationCity = $row."Destination City".Trim(); DestinationStateProvince = $row."Destination State".Trim(); DestinationPostalCode = $row."Destination Postal Code".Trim(); DestinationCountry = $row."Destination Country".Trim();
                BillToAccount = $null # Placeholder
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
            if ([string]::IsNullOrWhiteSpace($commodity.weight) -or ($commodity.weight -as [decimal]) -eq $null -or ([decimal]$commodity.weight -le 0)) { $commValid = $false; Write-Verbose "Skip Row ${rowNum}: Invalid weight '$($commodity.weight)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.length) -or ($commodity.length -as [decimal]) -eq $null -or ([decimal]$commodity.length -le 0)) { $commValid = $false; Write-Verbose "Skip Row ${rowNum}: Invalid length '$($commodity.length)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.width) -or ($commodity.width -as [decimal]) -eq $null -or ([decimal]$commodity.width -le 0)) { $commValid = $false; Write-Verbose "Skip Row ${rowNum}: Invalid width '$($commodity.width)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.height) -or ($commodity.height -as [decimal]) -eq $null -or ([decimal]$commodity.height -le 0)) { $commValid = $false; Write-Verbose "Skip Row ${rowNum}: Invalid height '$($commodity.height)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.pieces) -or ($commodity.pieces -as [int]) -eq $null -or ([int]$commodity.pieces -le 0)) { $commValid = $false; Write-Verbose "Skip Row ${rowNum}: Invalid pieces '$($commodity.pieces)'" }
            if ([string]::IsNullOrWhiteSpace($commodity.classification)) { Write-Verbose "Row ${rowNum}: Missing classification for commodity 1." }

            if ($commValid) {
                $normalizedEntry.Commodities += $commodity
            } else {
                $invalidRowCount++; continue # Skip row if essential commodity data is bad
            }

            # --- Basic Validation for Base Fields ---
            $baseValid = $true
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.OriginCity)) { $baseValid = $false }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.OriginStateProvince)) { $baseValid = $false }
            if ([string]::IsNullOrWhiteSpace($normalizedEntry.OriginPostalCode)) { $baseValid = $false }
            # Add other base field checks as needed...

            if (-not $baseValid) { Write-Warning "Skipping row $rowNum due to missing Origin/Dest fields."; $invalidRowCount++; continue }

            $normData.Add([PSCustomObject]$normalizedEntry)
        } # End foreach row

        if ($invalidRowCount -gt 0) { Write-Warning " -> Skipped $invalidRowCount Averitt rows." }
        Write-Host " -> OK: $($normData.Count) Averitt rows normalized." -ForegroundColor Green
        return $normData
    } catch { Write-Error "Error processing Averitt CSV '$CsvPath': $($_.Exception.Message)"; return $null }
}


# --- API Call Functions ---

function Invoke-AverittApi {
    param(
        [Parameter(Mandatory=$true)] [hashtable]$KeyData,
        [Parameter(Mandatory=$true)] [PSCustomObject]$ShipmentData # Normalized data object
    )

    $apiKey = $KeyData.APIKey
    $tariffNameForLog = if ($KeyData.TariffFileName) { $KeyData.TariffFileName } elseif ($KeyData.Name) {$KeyData.Name} else { "UnknownAverittTariff" }

    # --- Input Validation ---
    if ([string]::IsNullOrWhiteSpace($apiKey)) { Write-Warning "Averitt Skip ($tariffNameForLog): APIKey missing."; return $null }
    if ($null -eq $ShipmentData) { Write-Warning "Averitt Skip ($tariffNameForLog): ShipmentData null."; return $null }
    if ($null -eq $ShipmentData.Commodities -or $ShipmentData.Commodities.Count -eq 0) { Write-Warning "Averitt Skip ($tariffNameForLog): No commodities."; return $null }

    # --- Determine Required Account Number based on Payer ---
    # Attempt to get account number from ShipmentData first, then fallback to KeyData if needed
    # This assumes KeyData might contain a default account number for the tariff.
    $originAccountNum = if ($ShipmentData.PSObject.Properties.Name -contains 'OriginAccount' -and -not [string]::IsNullOrWhiteSpace($ShipmentData.OriginAccount)) { $ShipmentData.OriginAccount } elseif ($KeyData.ContainsKey('AccountNumber')) { $KeyData.AccountNumber } else { $null }
    $destAccountNum = if ($ShipmentData.PSObject.Properties.Name -contains 'DestinationAccount' -and -not [string]::IsNullOrWhiteSpace($ShipmentData.DestinationAccount)) { $ShipmentData.DestinationAccount } elseif ($KeyData.ContainsKey('AccountNumber')) { $KeyData.AccountNumber } else { $null }
    $billToAccountNum = if ($ShipmentData.PSObject.Properties.Name -contains 'BillToAccount' -and -not [string]::IsNullOrWhiteSpace($ShipmentData.BillToAccount)) { $ShipmentData.BillToAccount } elseif ($KeyData.ContainsKey('AccountNumber')) { $KeyData.AccountNumber } else { $null }

    # Validate required account based on payer
    $payer = $ShipmentData.PaymentPayer
    if ($payer -eq 'Shipper' -and [string]::IsNullOrWhiteSpace($originAccountNum)) { Write-Warning "Averitt Skip ($tariffNameForLog): Payer is Shipper but Origin Account number is missing."; return $null }
    if ($payer -eq 'Consignee' -and [string]::IsNullOrWhiteSpace($destAccountNum)) { Write-Warning "Averitt Skip ($tariffNameForLog): Payer is Consignee but Destination Account number is missing."; return $null }
    if ($payer -eq 'ThirdParty' -and [string]::IsNullOrWhiteSpace($billToAccountNum)) { Write-Warning "Averitt Skip ($tariffNameForLog): Payer is ThirdParty but BillTo Account number is missing."; return $null }


    # --- Construct Payload ---
    $payloadObject = [ordered]@{
        service = @{ level = $ShipmentData.ServiceLevel }
        payment = @{ terms = $ShipmentData.PaymentTerms; payer = $payer } # Use determined payer
        transit = @{ pickupDate = $ShipmentData.PickupDate }
        commodities = @() # Initialize
        origin = @{
            account = $originAccountNum # Use determined account
            city = $ShipmentData.OriginCity
            stateProvince = $ShipmentData.OriginStateProvince
            postalCode = $ShipmentData.OriginPostalCode
            country = $ShipmentData.OriginCountry
        }
        destination = @{
            account = $destAccountNum # Use determined account
            city = $ShipmentData.DestinationCity
            stateProvince = $ShipmentData.DestinationStateProvince
            postalCode = $ShipmentData.DestinationPostalCode
            country = $ShipmentData.DestinationCountry
        }
        billTo = @{ # Required by API
            account = $billToAccountNum # Use determined account
            name = $ShipmentData.BillToName
            address = $ShipmentData.BillToAddress
            city = $ShipmentData.BillToCity
            stateProvince = $ShipmentData.BillToStateProvince
            postalCode = $ShipmentData.BillToPostalCode
            country = $ShipmentData.BillToCountry
        }
    }
    foreach ($comm in $ShipmentData.Commodities) {
        $apiComm = [ordered]@{ classification = [string]$comm.classification; weight = [string]$comm.weight; length = [string]$comm.length; width = [string]$comm.width; height = [string]$comm.height; pieces = [string]$comm.pieces; packagingType = [string]$comm.packagingType; description = [string]$comm.description; stackable = [string]$comm.stackable }
        $payloadObject.commodities += $apiComm
    }
    if ($null -ne $ShipmentData.Accessorials) { $payloadObject.accessorials = $ShipmentData.Accessorials; if (!$payloadObject.accessorials.ContainsKey('hazardousContact')) { $payloadObject.accessorials.hazardousContact = $null }; if (!$payloadObject.accessorials.ContainsKey('cod')) { $payloadObject.accessorials.cod = $null }; if (!$payloadObject.accessorials.ContainsKey('insuranceDetails')) { $payloadObject.accessorials.insuranceDetails = $null }; if (!$payloadObject.accessorials.ContainsKey('sortAndSegregateDetails')) { $payloadObject.accessorials.sortAndSegregateDetails = $null }; if (!$payloadObject.accessorials.ContainsKey('markDetails')) { $payloadObject.accessorials.markDetails = $null } } else { $payloadObject.accessorials = @{ codes = @(); hazardousContact=$null; cod=$null; insuranceDetails=$null; sortAndSegregateDetails=$null; markDetails=$null } }

    # --- Convert to JSON ---
    $payloadJson = $payloadObject | ConvertTo-Json -Depth 10
    Write-Verbose "Averitt Payload ($tariffNameForLog): $payloadJson"

    # --- API Call ---
    $apiUrl = $script:averittApiUri
    if ([string]::IsNullOrWhiteSpace($apiUrl)) { Write-Error "Averitt API URI not configured."; return $null }
    $uri = "{0}?api_key={1}" -f $apiUrl, $apiKey
    $headers = @{ 'Content-Type' = 'application/json'; 'Accept' = 'application/json' }

    Write-Verbose "Calling Averitt API: $tariffNameForLog"
    try {
        $response = Invoke-RestMethod -Uri $uri -Method Post -Headers $headers -Body $payloadJson -ErrorAction Stop
        Write-Verbose "Averitt Raw Response ($tariffNameForLog): $($response | ConvertTo-Json -Depth 5)"

        # --- Process Response ---
        if ($response -ne $null -and $response.PSObject.Properties.Name -contains 'messageStatus' -and $response.messageStatus -ne $null) {
            if ($response.messageStatus.status -in @('PASS', 'WARNING')) {
                 $totalChargeValue = $null
                 if ($response.PSObject.Properties.Name -contains 'shipmentInfo' -and $response.shipmentInfo -ne $null -and $response.shipmentInfo.PSObject.Properties.Name -contains 'totalCharge') {
                      $totalChargeValue = $response.shipmentInfo.totalCharge
                 }

                 if ($totalChargeValue -ne $null) {
                     try {
                         $cleanedRate = $totalChargeValue -replace '[$,]'; $decimalRate = [decimal]$cleanedRate
                         if ($decimalRate -lt 0) { throw "Negative rate."}
                         Write-Verbose "Averitt OK ($tariffNameForLog): Rate = $decimalRate"
                         return $decimalRate
                     } catch { Write-Warning "Averitt Convert Fail ($tariffNameForLog): Cannot convert '$totalChargeValue'. Error: $($_.Exception.Message)"; return $null }
                 } else { Write-Warning "Averitt Rate Not Found ($tariffNameForLog): 'shipmentInfo.totalCharge' missing/null."; return $null }
            } else { Write-Warning "Averitt API Fail ($tariffNameForLog): Status=$($response.messageStatus.status), Code=$($response.messageStatus.code), Msg=$($response.messageStatus.message)"; return $null }
        } else { Write-Warning "Averitt API Error ($tariffNameForLog): Response invalid or missing 'messageStatus'."; return $null }
    } catch {
        $errMsg = $_.Exception.Message; $statusCode = "N/A"; $eBody = "N/A"
        if ($_.Exception.Response) { try {$statusCode = $_.Exception.Response.StatusCode.value__} catch{}; try { $stream = $_.Exception.Response.GetResponseStream(); $reader = New-Object System.IO.StreamReader($stream); $eBody = $reader.ReadToEnd(); $reader.Close(); $stream.Close() } catch {$eBody="(Err reading resp body)"} }
        $truncatedBody = if ($eBody.Length -gt 500) { $eBody.Substring(0, 500) + "..." } else { $eBody }
        $fullErrMsg = "Averitt Invoke FAIL ($tariffNameForLog): Error: $errMsg (HTTP $statusCode) Resp: $truncatedBody"; Write-Warning $fullErrMsg; return $null
    }
}

Write-Verbose "TMS Averitt Helper Functions loaded."