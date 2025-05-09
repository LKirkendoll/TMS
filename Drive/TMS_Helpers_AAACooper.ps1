# TMS_Helpers_AAACooper.ps1
# Description: Contains helper functions specific to AAA Cooper Transportation operations,
#              including data normalization and API interaction.
#              This file should be dot-sourced by the main script(s) after TMS_Config.ps1.

# Assumes config variables like $script:aaaCooperApiUri are available from TMS_Config.ps1
# Assumes general helper functions are available.

# --- Data Normalization Functions ---

function Load-And-Normalize-AAACooperData {
    param(
        [Parameter(Mandatory=$true)][string]$CsvPath
    )
    Write-Host "`nLoading AAA Cooper data: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan

    # Define required columns based on AAA Cooper API needs
    # Base shipment details
    $reqBaseCols = @(
        "Origin City", "Origin State", "Origin Postal Code",
        "Destination City", "Destination State", "Destination Postal Code"
    )
    # Commodity details (assuming mapping from common "Total" columns for the first/primary commodity line)
    $reqCommCols = @(
        "Total Weight", "Freight Class 1", "Total Units"
    )
    # Optional commodity details
    $optCommCols = @(
        "Total Length", "Total Width", "Total Height", # For dimensions
        "NMFC Item 1", "NMFC Sub 1", "Handling Unit Type 1", "Hazmat 1"
    )

    try {
        if (-not (Test-Path -Path $CsvPath -PathType Leaf)) {
            Write-Error "CSV file not found at '$CsvPath'."
            return $null
        }
        $rawData = Import-Csv -Path $CsvPath -Delimiter ',' -ErrorAction Stop
        Write-Host " -> Rows read from CSV: $($rawData.Count)." -ForegroundColor Gray
        if ($rawData.Count -eq 0) { Write-Warning "CSV empty."; return @() }

        $headers = $rawData[0].PSObject.Properties.Name
        $missingBase = $reqBaseCols | Where-Object { $_ -notin $headers }
        if ($missingBase.Count -gt 0) { Write-Error "CSV missing required AAA Cooper base columns: $($missingBase -join ', ')"; return $null }
        $missingComm = $reqCommCols | Where-Object { $_ -notin $headers }
        if ($missingComm.Count -gt 0) { Write-Error "CSV missing required AAA Cooper commodity columns: $($missingComm -join ', ')"; return $null }

        $normData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing AAA Cooper data..." -ForegroundColor Gray
        $invalidCount = 0; $csvRowNum = 1

        foreach ($row in $rawData) {
            $currentDataRowForMessage = $csvRowNum + 1
            $skipRow = $false

            # Basic Shipment Details
            $originCity = $row."Origin City".Trim()
            $originState = $row."Origin State".Trim()
            $originZip = $row."Origin Postal Code".Trim()
            $originCountry = if ($headers -contains "Origin Country") { $row."Origin Country".Trim() } else { "USA" } # Default to USA

            $destCity = $row."Destination City".Trim()
            $destState = $row."Destination State".Trim()
            $destZip = $row."Destination Postal Code".Trim()
            $destCountry = if ($headers -contains "Destination Country") { $row."Destination Country".Trim() } else { "USA" } # Default to USA

            # Validate base details
            if ([string]::IsNullOrWhiteSpace($originCity)) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Missing Origin City" }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($originState) -or $originState.Length -ne 2)) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Invalid Origin State" }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($originZip) -or $originZip.Length -gt 6)) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Invalid Origin Zip" }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($destCity))) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Missing Destination City" }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($destState) -or $destState.Length -ne 2)) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Invalid Destination State" }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($destZip) -or $destZip.Length -gt 6)) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Invalid Destination Zip" }


            # Commodity Details (mapping "Total" fields to the first commodity line)
            $commWeightStr = $row."Total Weight".Trim()
            $commClassStr = $row."Freight Class 1".Trim()
            $commHandlingUnitsStr = $row."Total Units".Trim()

            $commWeight = $null; $commClass = $null; $commHandlingUnits = $null

            if (-not $skipRow) {
                try { $commWeight = [int]$commWeightStr; if ($commWeight -lt 2) { throw "Weight must be >= 2."} } catch { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Invalid Total Weight '$($commWeightStr)'" }
            }
            if (-not $skipRow) {
                if ([string]::IsNullOrWhiteSpace($commClassStr) -or -not ($commClassStr -match '^\d+(\.\d+)?$') ) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Invalid Freight Class 1 '$($commClassStr)'" } else { $commClass = $commClassStr }
            }
            if (-not $skipRow) {
                try { $commHandlingUnits = [int]$commHandlingUnitsStr; if ($commHandlingUnits -le 0) { throw "Handling units must be > 0."} } catch { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Invalid Total Units '$($commHandlingUnitsStr)'" }
            }

            if ($skipRow) { $invalidCount++; $csvRowNum++; continue }

            # Optional commodity fields
            $commNMFC = if ($headers -contains "NMFC Item 1") { $row."NMFC Item 1".Trim() } else { $null }
            $commNMFCSub = if ($headers -contains "NMFC Sub 1") { $row."NMFC Sub 1".Trim() } else { $null }
            $commHandlingUnitType = if ($headers -contains "Handling Unit Type 1") { $row."Handling Unit Type 1".Trim() } else { "Pallets" } # Default
            $commHazMat = if ($headers -contains "Hazmat 1" -and ($row."Hazmat 1".Trim() -eq 'TRUE' -or $row."Hazmat 1".Trim() -eq 'X')) { "X" } else { $null }
            $commLength = if ($headers -contains "Total Length") { try { [int]($row."Total Length".Trim()) } catch { $null } } else { $null }
            $commWidth = if ($headers -contains "Total Width") { try { [int]($row."Total Width".Trim()) } catch { $null } } else { $null }
            $commHeight = if ($headers -contains "Total Height") { try { [int]($row."Total Height".Trim()) } catch { $null } } else { $null }


            $commodityLine = [PSCustomObject]@{
                Weight = $commWeight
                Class = $commClass
                NMFC = $commNMFC
                NMFCSub = $commNMFCSub
                HandlingUnits = $commHandlingUnits
                HandlingUnitType = $commHandlingUnitType
                HazMat = $commHazMat
                CubeUnit = if ($commLength -and $commWidth -and $commHeight) { "IN" } else { $null } # Only if all dims present
                Length = $commLength
                Width = $commWidth
                Height = $commHeight
            }

            $normalizedEntry = [PSCustomObject]@{
                OriginCity = $originCity
                OriginState = $originState
                OriginZip = $originZip
                OriginCountryCode = $originCountry
                DestinationCity = $destCity
                DestinationState = $destState
                DestinationZip = $destZip
                DestinCountryCode = $destCountry # Corrected typo from Destin to Destination
                BillDate = (Get-Date -Format 'MMddyyyy') # API format MMDDYY or MMDDYYYY
                PrePaidCollect = "P" # Default to Prepaid, can be overridden by KeyData or specific logic
                RateEstimateRequestLine = @($commodityLine)
                # Store original values if needed for other reports
                'Total Weight' = $commWeight
                'Freight Class 1' = $commClass
                'Total Units' = $commHandlingUnits
            }
            $normData.Add($normalizedEntry)
            $csvRowNum++
        }

        if ($invalidCount -gt 0) { Write-Warning " -> Skipped $invalidCount AAA Cooper rows due to validation errors." }
        Write-Host " -> OK: $($normData.Count) AAA Cooper rows normalized." -ForegroundColor Green
        return $normData

    } catch {
        Write-Error "Error processing AAA Cooper CSV '$CsvPath': $($_.Exception.Message)"
        return $null
    }
}


# --- API Call Functions ---

function Invoke-AAACooperApi {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$KeyData, # Expects 'APIToken', 'CustomerNumber', 'WhoAmI', optionally 'PrePaidCollect'
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$ShipmentData # Normalized data object from Load-And-Normalize-AAACooperData
    )

    $tariffNameForLog = if ($KeyData.ContainsKey('TariffFileName')) { $KeyData.TariffFileName } elseif ($KeyData.ContainsKey('Name')) {$KeyData.Name} else { "AAACooperTariff" }
    Write-Verbose "Attempting AAA Cooper API call for Tariff: $tariffNameForLog"

    # --- Extract credentials and required fields from KeyData ---
    $apiToken = $KeyData.APIToken
    $customerNumber = $KeyData.CustomerNumber
    $whoAmI = $KeyData.WhoAmI
    $prePaidCollect = if ($KeyData.ContainsKey('PrePaidCollect')) { $KeyData.PrePaidCollect } else { $ShipmentData.PrePaidCollect } # Use from ShipmentData if not in KeyData

    # --- Validate KeyData ---
    if ([string]::IsNullOrWhiteSpace($apiToken)) { Write-Warning "AAA Cooper Skip ($tariffNameForLog): APIToken missing from KeyData."; return $null }
    if ([string]::IsNullOrWhiteSpace($customerNumber)) { Write-Warning "AAA Cooper Skip ($tariffNameForLog): CustomerNumber missing from KeyData."; return $null }
    if ([string]::IsNullOrWhiteSpace($whoAmI)) { Write-Warning "AAA Cooper Skip ($tariffNameForLog): WhoAmI missing from KeyData."; return $null }

    # --- Validate ShipmentData ---
    if ($null -eq $ShipmentData) { Write-Warning "AAA Cooper Skip ($tariffNameForLog): ShipmentData object is null."; return $null }
    if ($null -eq $ShipmentData.RateEstimateRequestLine -or $ShipmentData.RateEstimateRequestLine.Count -eq 0) {
        Write-Warning "AAA Cooper Skip ($tariffNameForLog): No commodity lines (RateEstimateRequestLine) in ShipmentData."
        return $null
    }

    # Helper to escape XML special characters
    function Escape-XmlValue ($value) {
        if ($null -eq $value) { return '' }
        return [System.Security.SecurityElement]::Escape($value.ToString())
    }

    # --- Construct SOAP XML Payload ---
    $xmlWriterSettings = New-Object System.Xml.XmlWriterSettings
    $xmlWriterSettings.Indent = $true
    $xmlWriterSettings.OmitXmlDeclaration = $true # SOAP envelope will have its own

    $stringBuilder = New-Object System.Text.StringBuilder
    $xmlWriter = [System.Xml.XmlWriter]::Create($stringBuilder, $xmlWriterSettings)

    $soapNs = "http://schemas.xmlsoap.org/soap/envelope/"
    $tempUriNs = "http://tempuri.org/wsGenRateEstimate/" # Namespace from WSDL/Sample

    $xmlWriter.WriteStartElement("soap", "Envelope", $soapNs)
    $xmlWriter.WriteAttributeString("xmlns", "xsi", $null, "http://www.w3.org/2001/XMLSchema-instance")
    $xmlWriter.WriteAttributeString("xmlns", "xsd", $null, "http://www.w3.org/2001/XMLSchema")

    $xmlWriter.WriteStartElement("soap", "Body", $soapNs)
    $xmlWriter.WriteStartElement("RateEstimateRequestVO", $tempUriNs)

    # Required elements
    $xmlWriter.WriteElementString("Token", $tempUriNs, (Escape-XmlValue $apiToken))
    $xmlWriter.WriteElementString("CustomerNumber", $tempUriNs, (Escape-XmlValue $customerNumber))
    $xmlWriter.WriteElementString("OriginCity", $tempUriNs, (Escape-XmlValue $ShipmentData.OriginCity))
    $xmlWriter.WriteElementString("OriginState", $tempUriNs, (Escape-XmlValue $ShipmentData.OriginState))
    $xmlWriter.WriteElementString("OriginZip", $tempUriNs, (Escape-XmlValue $ShipmentData.OriginZip))
    $xmlWriter.WriteElementString("OriginCountryCode", $tempUriNs, (Escape-XmlValue $ShipmentData.OriginCountryCode))
    $xmlWriter.WriteElementString("DestinationCity", $tempUriNs, (Escape-XmlValue $ShipmentData.DestinationCity))
    $xmlWriter.WriteElementString("DestinationState", $tempUriNs, (Escape-XmlValue $ShipmentData.DestinationState))
    $xmlWriter.WriteElementString("DestinationZip", $tempUriNs, (Escape-XmlValue $ShipmentData.DestinationZip))
    $xmlWriter.WriteElementString("DestinCountryCode", $tempUriNs, (Escape-XmlValue $ShipmentData.DestinCountryCode)) # API Doc uses "DestinCountryCode"
    $xmlWriter.WriteElementString("WhoAmI", $tempUriNs, (Escape-XmlValue $whoAmI))
    $xmlWriter.WriteElementString("BillDate", $tempUriNs, (Escape-XmlValue $ShipmentData.BillDate))
    $xmlWriter.WriteElementString("PrePaidCollect", $tempUriNs, (Escape-XmlValue $prePaidCollect))

    # Optional: TotalPalletCount (Example, if you have this data)
    # if ($ShipmentData.PSObject.Properties.Name -contains 'TotalPalletCount' -and $ShipmentData.TotalPalletCount) {
    #     $xmlWriter.WriteElementString("TotalPalletCount", $tempUriNs, (Escape-XmlValue $ShipmentData.TotalPalletCount))
    # }

    # RateEstimateRequestLine (Commodities)
    if ($ShipmentData.RateEstimateRequestLine.Count -gt 0) {
        foreach ($lineItem in $ShipmentData.RateEstimateRequestLine) {
            $xmlWriter.WriteStartElement("RateEstimateRequestLine", $tempUriNs)
            $xmlWriter.WriteElementString("Weight", $tempUriNs, (Escape-XmlValue $lineItem.Weight))
            $xmlWriter.WriteElementString("Class", $tempUriNs, (Escape-XmlValue $lineItem.Class))
            if (-not [string]::IsNullOrWhiteSpace($lineItem.NMFC)) { $xmlWriter.WriteElementString("NMFC", $tempUriNs, (Escape-XmlValue $lineItem.NMFC)) }
            if (-not [string]::IsNullOrWhiteSpace($lineItem.NMFCSub)) { $xmlWriter.WriteElementString("NMFCSub", $tempUriNs, (Escape-XmlValue $lineItem.NMFCSub)) }
            if ($lineItem.HandlingUnits) { $xmlWriter.WriteElementString("HandlingUnits", $tempUriNs, (Escape-XmlValue $lineItem.HandlingUnits)) }
            if (-not [string]::IsNullOrWhiteSpace($lineItem.HandlingUnitType)) { $xmlWriter.WriteElementString("HandlingUnitType", $tempUriNs, (Escape-XmlValue $lineItem.HandlingUnitType)) }
            if (-not [string]::IsNullOrWhiteSpace($lineItem.HazMat)) { $xmlWriter.WriteElementString("HazMat", $tempUriNs, (Escape-XmlValue $lineItem.HazMat)) }
            if (-not [string]::IsNullOrWhiteSpace($lineItem.CubeUnit)) { $xmlWriter.WriteElementString("CubeUnit", $tempUriNs, (Escape-XmlValue $lineItem.CubeUnit)) }
            if ($lineItem.Length) { $xmlWriter.WriteElementString("Length", $tempUriNs, (Escape-XmlValue $lineItem.Length)) } # Length required if > 96
            if ($lineItem.Width) { $xmlWriter.WriteElementString("Width", $tempUriNs, (Escape-XmlValue $lineItem.Width)) }
            if ($lineItem.Height) { $xmlWriter.WriteElementString("Height", $tempUriNs, (Escape-XmlValue $lineItem.Height)) }
            $xmlWriter.WriteEndElement() # RateEstimateRequestLine
        }
    }

    # Optional: AccLine (Accessorials) - Example structure
    # if ($ShipmentData.Accessorials -and $ShipmentData.Accessorials.Count -gt 0) {
    #     foreach ($accCode in $ShipmentData.Accessorials) {
    #         $xmlWriter.WriteStartElement("AccLine", $tempUriNs)
    #         $xmlWriter.WriteElementString("AccCode", $tempUriNs, (Escape-XmlValue $accCode))
    #         $xmlWriter.WriteEndElement() # AccLine
    #     }
    # }

    $xmlWriter.WriteEndElement() # RateEstimateRequestVO
    $xmlWriter.WriteEndElement() # soap:Body
    $xmlWriter.WriteEndElement() # soap:Envelope
    $xmlWriter.Flush()
    $xmlWriter.Close()

    $soapRequestXml = $stringBuilder.ToString()
    Write-Verbose "AAA Cooper SOAP Request XML for Tariff '$tariffNameForLog':`n$soapRequestXml"

    # --- API Call ---
    $apiUrl = $script:aaaCooperApiUri
    if ([string]::IsNullOrWhiteSpace($apiUrl)) {
        Write-Error "AAA Cooper API URI not configured in TMS_Config.ps1."
        return $null
    }
    $headers = @{
        "Content-Type" = "text/xml;charset=UTF-8"
        "SOAPAction"   = "http://tempuri.org/wsGenRateEstimate/RateEstimateRequestVO" # Often required, check WSDL or examples
    }

    try {
        $ProgressPreference = 'SilentlyContinue'
        $response = Invoke-WebRequest -Uri $apiUrl -Method Post -Headers $headers -Body $soapRequestXml -UseBasicParsing -ErrorAction Stop
        [xml]$responseXml = $response.Content
        Write-Verbose "AAA Cooper Raw Response XML for Tariff '$tariffNameForLog':`n$($responseXml.OuterXml)"

        # --- Process Response ---
        $nsManager = New-Object System.Xml.XmlNamespaceManager($responseXml.NameTable)
        $nsManager.AddNamespace("soap", $soapNs)
        $nsManager.AddNamespace("ns1", $tempUriNs) # Namespace used in sample response

        $faultNode = $responseXml.SelectSingleNode("/soap:Envelope/soap:Body/soap:Fault", $nsManager)
        if ($faultNode) {
            $faultString = $faultNode.SelectSingleNode("faultstring", $nsManager).InnerText
            $faultCode = $faultNode.SelectSingleNode("faultcode", $nsManager).InnerText
            Write-Warning "AAA Cooper API returned SOAP Fault for Tariff '$tariffNameForLog': Code='$faultCode', String='$faultString'"
            return $null
        }

        $rateResponseNode = $responseXml.SelectSingleNode("/soap:Envelope/soap:Body/ns1:RateEstimateResponseVO", $nsManager)
        if (-not $rateResponseNode) {
            Write-Warning "AAA Cooper API Response Error ($tariffNameForLog): RateEstimateResponseVO node not found."
            return $null
        }

        $errorMessageNode = $rateResponseNode.SelectSingleNode("ns1:ErrorMessage", $nsManager)
        if ($errorMessageNode -and -not [string]::IsNullOrWhiteSpace($errorMessageNode.InnerText)) {
            Write-Warning "AAA Cooper API Error ($tariffNameForLog): $($errorMessageNode.InnerText)"
            return $null
        }

        $totalChargesNode = $rateResponseNode.SelectSingleNode("ns1:TotalCharges", $nsManager)
        if ($totalChargesNode -and -not [string]::IsNullOrWhiteSpace($totalChargesNode.InnerText)) {
            $totalChargeValue = $totalChargesNode.InnerText
            try {
                $cleanedRate = $totalChargeValue -replace '[$,]'
                $decimalRate = [decimal]$cleanedRate
                if ($decimalRate -lt 0) { throw "Negative rate received from API." }
                Write-Verbose "AAA Cooper API Call OK ($tariffNameForLog): Rate = $decimalRate"
                return $decimalRate
            } catch {
                Write-Warning "AAA Cooper Rate Conversion Fail ($tariffNameForLog): Cannot convert rate '$totalChargeValue' to decimal. Error: $($_.Exception.Message)"
                return $null
            }
        } else {
            Write-Warning "AAA Cooper Rate Not Found in Response ($tariffNameForLog): 'TotalCharges' field is missing, null, or empty in API response."
            return $null
        }
    } catch {
        $errMsg = $_.Exception.Message
        $statusCode = "N/A"
        $eBody = "N/A"
        if ($_.Exception.Response) {
            try { $statusCode = $_.Exception.Response.StatusCode.value__ } catch { }
            try {
                $stream = $_.Exception.Response.GetResponseStream()
                $reader = New-Object System.IO.StreamReader($stream)
                $eBody = $reader.ReadToEnd()
                $reader.Close(); $stream.Close()
            } catch { $eBody = "(Error reading response body: $($_.Exception.Message))" }
        }
        $truncatedBody = if ($eBody.Length -gt 500) { $eBody.Substring(0, 500) + "..." } else { $eBody }
        $fullErrMsg = "AAA Cooper API Invoke-WebRequest FAILED ($tariffNameForLog): Error: $errMsg (HTTP Status: $statusCode) Response Body: $truncatedBody"
        Write-Warning $fullErrMsg
        return $null
    }
}

Write-Verbose "TMS AAA Cooper Helper Functions loaded."
