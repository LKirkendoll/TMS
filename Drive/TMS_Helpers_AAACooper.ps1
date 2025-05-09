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

    $reqBaseCols = @(
        "Origin City", "Origin State", "Origin Postal Code",
        "Destination City", "Destination State", "Destination Postal Code"
    )
    $reqCommCols = @(
        "Total Weight", "Freight Class 1", "Total Units"
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

            $originCity = $row."Origin City".Trim()
            $originState = $row."Origin State".Trim()
            $originZip = $row."Origin Postal Code".Trim()
            $originCountry = if ($headers -contains "Origin Country") { $row."Origin Country".Trim() } else { "USA" }

            $destCity = $row."Destination City".Trim()
            $destState = $row."Destination State".Trim()
            $destZip = $row."Destination Postal Code".Trim()
            $destCountry = if ($headers -contains "Destination Country") { $row."Destination Country".Trim() } else { "USA" }

            if ([string]::IsNullOrWhiteSpace($originCity) ) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Missing Origin City" } 
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($originState) -or $originState.Length -ne 2)) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Invalid Origin State ('$($originState)')" }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($originZip) -or $originZip.Length -lt 5 -or $originZip.Length -gt 6 )) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Invalid Origin Zip ('$($originZip)')" }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($destCity))) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Missing Destination City" }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($destState) -or $destState.Length -ne 2)) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Invalid Destination State ('$($destState)')" }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($destZip) -or $destZip.Length -lt 5 -or $destZip.Length -gt 6)) { $skipRow = $true; Write-Verbose "Skip AAA Cooper Row ${currentDataRowForMessage}: Invalid Destination Zip ('$($destZip)')" }

            $commWeightStr = $row."Total Weight".Trim()
            $commClassStr = $row."Freight Class 1".Trim()
            $commHandlingUnitsStr = $row."Total Units".Trim()
            
            $commWeight = $null; $commClass = $null; $commHandlingUnits = $null
            $isWeightValid = $false

            if (-not $skipRow) {
                # Validate Weight first
                if (-not [string]::IsNullOrWhiteSpace($commWeightStr)) {
                    try {
                        $tempWeight = [int]$commWeightStr
                        if ($tempWeight -ge 2) { # As per PDF, weight must be >= 2
                            $commWeight = $tempWeight
                            $isWeightValid = $true
                        } else {
                            throw "Weight must be >= 2."
                        }
                    } catch {
                        Write-Verbose "Skip Row $($currentDataRowForMessage): Invalid Total Weight '$($commWeightStr)'. Error: $($_.Exception.Message)"
                        $skipRow = $true
                    }
                } else {
                    # Weight is blank, which might be okay if class is also blank (empty line)
                    # or might be an error if class is provided. For now, we'll let class validation handle it.
                    $isWeightValid = $false 
                }
            }

            if (-not $skipRow) {
                # Validate Class - It's required if Weight is valid and present.
                if ($isWeightValid) { # If weight is valid, class MUST be valid
                    if ([string]::IsNullOrWhiteSpace($commClassStr) -or -not ($commClassStr -match '^\d+(\.\d+)?$')) {
                        Write-Verbose "Skip Row $($currentDataRowForMessage): Class ('$commClassStr') is missing or invalid when Weight ('$commWeightStr') is present and valid."
                        $skipRow = $true
                    } else {
                        $commClass = $commClassStr
                    }
                } elseif (-not [string]::IsNullOrWhiteSpace($commClassStr)) { 
                    # If weight is NOT valid/present, but class IS provided, validate class format anyway.
                    # The API will likely reject an item with class but no weight.
                    if (-not ($commClassStr -match '^\d+(\.\d+)?$')) {
                        Write-Verbose "Skip Row $($currentDataRowForMessage): Invalid Class format ('$commClassStr') (Weight was not valid/present)."
                        $skipRow = $true 
                    } else {
                         $commClass = $commClassStr # Store it if format is okay
                    }
                }
                # If both weight and class are blank/invalid, the row will likely be skipped or $commWeight/$commClass will be $null.
            }

            if (-not $skipRow) {
                # Validate HandlingUnits - Required if weight & class are present.
                # For simplicity, we'll make it required if weight is valid, as items usually have handling units.
                if ($isWeightValid) { # Only enforce handling units if weight is valid
                    if ([string]::IsNullOrWhiteSpace($commHandlingUnitsStr)) {
                        Write-Verbose "Skip Row $($currentDataRowForMessage): Handling Units are missing when Weight is present."
                        $skipRow = $true
                    } else {
                        try {
                            $commHandlingUnits = [int]$commHandlingUnitsStr
                            if ($commHandlingUnits -le 0) { throw "Handling units must be > 0." }
                        } catch {
                            Write-Verbose "Skip Row $($currentDataRowForMessage): Invalid Total Units '$($commHandlingUnitsStr)'. Error: $($_.Exception.Message)"
                            $skipRow = $true
                        }
                    }
                } elseif (-not [string]::IsNullOrWhiteSpace($commHandlingUnitsStr)) { # If weight is not valid, but HU is provided, try to parse
                     try { $commHandlingUnits = [int]$commHandlingUnitsStr } catch { Write-Verbose "Warning Row $($currentDataRowForMessage): Could not parse Handling Units '$commHandlingUnitsStr', but Weight was also invalid."}
                }
            }
            
            if ($skipRow) { $invalidCount++; $csvRowNum++; continue } 
            
            # If we reach here, and $isWeightValid is true, then $commWeight AND $commClass should be valid.
            # If $isWeightValid is false, $commWeight is $null. $commClass might be $null or have a value.
            # The API call itself (Invoke-AAACooperApi) has a final check.

            $rawHandlingUnitType = if ($headers -contains "Handling Unit Type 1" -and -not [string]::IsNullOrWhiteSpace($row."Handling Unit Type 1")) { $row."Handling Unit Type 1".Trim().ToUpperInvariant() } else { "PALLETS" } 
            $commHandlingUnitType = if ($rawHandlingUnitType -in ("PLT", "PAT", "PALLET", "PALLETS")) { "Pallets" } else { $rawHandlingUnitType } 
            
            $commHazMat = if ($headers -contains "Hazmat 1" -and ($row."Hazmat 1".Trim().ToUpperInvariant() -in @('TRUE', 'X', 'Y'))) { "X" } else { $null }
            $commLength = if ($headers -contains "Total Length") { try { [int]($row."Total Length".Trim()) } catch { $null } } else { $null }
            $commWidth = if ($headers -contains "Total Width") { try { [int]($row."Total Width".Trim()) } catch { $null } } else { $null }
            $commHeight = if ($headers -contains "Total Height") { try { [int]($row."Total Height".Trim()) } catch { $null } } else { $null }

            $commodityLine = [PSCustomObject]@{
                Weight = $commWeight 
                Class = $commClass   
                NMFC = if ($headers -contains "NMFC Item 1") { $row."NMFC Item 1".Trim() } else { $null }
                NMFCSub = if ($headers -contains "NMFC Sub 1") { $row."NMFC Sub 1".Trim() } else { $null }
                HandlingUnits = $commHandlingUnits 
                HandlingUnitType = $commHandlingUnitType 
                HazMat = $commHazMat
                CubeU = if ($commLength -and $commWidth -and $commHeight -and $commLength -gt 0 -and $commWidth -gt 0 -and $commHeight -gt 0) { "IN" } else { $null } 
                Length = if($commLength -and $commLength -gt 0) {$commLength} else {$null}
                Width = if($commWidth -and $commWidth -gt 0) {$commWidth} else {$null}
                Height = if($commHeight -and $commHeight -gt 0) {$commHeight} else {$null}
            }
            
            $totalPalletCountForPayload = if ($commHandlingUnitType -eq "Pallets" -and $commHandlingUnits -gt 0) { $commHandlingUnits } else { 0 } 
            $primaryAccCode = if ($commHandlingUnitType -eq "Pallets" -and $commHandlingUnits -gt 0) { "PALET" } else { $null } 

            $normalizedEntry = [PSCustomObject]@{
                OriginCity = $originCity; OriginState = $originState; OriginZip = $originZip; OriginCountryCode = $originCountry
                DestinationCity = $destCity; DestinationState = $destState; DestinationZip = $destZip; DestinCountryCode = $destCountry 
                BillDate = (Get-Date -Format 'MMddyyyy'); WhoAmI = "S" 
                PrePaidCollect = "P" 
                TotalPalletCountForPayload = $totalPalletCountForPayload 
                PrimaryAccCode = $primaryAccCode 
                RateEstimateRequestLine = @($commodityLine) 
            }
            $normData.Add($normalizedEntry)
            $csvRowNum++
        }
        if ($invalidCount -gt 0) { Write-Warning " -> Skipped $invalidCount AAA Cooper rows due to missing/invalid Weight, Class, or HandlingUnits." }
        Write-Host " -> OK: $($normData.Count) AAA Cooper rows normalized." -ForegroundColor Green
        return $normData
    } catch {
        Write-Error "Error processing AAA Cooper CSV '$CsvPath': $($_.Exception.Message)"
        return $null
    }
}

function Invoke-AAACooperApi {
    param(
        [Parameter(Mandatory=$true)] [hashtable]$KeyData, 
        [Parameter(Mandatory=$true)] [PSCustomObject]$ShipmentData 
    )
    $tariffNameForLog = if ($KeyData.ContainsKey('Name')) { $KeyData.Name } else { $KeyData.TariffFileName | Split-Path -LeafBase }
    Write-Verbose "Attempting AAA Cooper API call for Tariff: $tariffNameForLog"

    $apiToken = $KeyData.APIToken
    $customerNumber = $KeyData.CustomerNumber 
    $apiWhoAmI = $KeyData.WhoAmI 
    $prePaidCollect = "P" 
    Write-Verbose "Invoke-AAACooperApi: Forcing PrePaidCollect to '$prePaidCollect'."

    if ([string]::IsNullOrWhiteSpace($apiToken)) { Write-Warning "${tariffNameForLog}: APIToken missing."; return $null }
    if ([string]::IsNullOrWhiteSpace($apiWhoAmI)) { Write-Warning "${tariffNameForLog}: WhoAmI from KeyData is missing."; return $null }
    if ($null -eq $ShipmentData) { Write-Warning "${tariffNameForLog}: ShipmentData object is null."; return $null }
    
    if (-not (Get-Command Escape-Xml -EA SilentlyContinue)) {
        function Escape-Xml ($value) { if ($null -eq $value) { return '' }; return [System.Security.SecurityElement]::Escape($value.ToString()) }
        Write-Warning "Local Escape-Xml used in Invoke-AAACooperApi."
    }

    $xmlWriterSettings = New-Object System.Xml.XmlWriterSettings
    $xmlWriterSettings.Indent = $true
    $xmlWriterSettings.OmitXmlDeclaration = $true 
    $xmlWriterSettings.Encoding = New-Object System.Text.UTF8Encoding($false) 

    $stringBuilder = New-Object System.Text.StringBuilder
    $xmlWriter = [System.Xml.XmlWriter]::Create($stringBuilder, $xmlWriterSettings)

    $soapEnvNs = "http://schemas.xmlsoap.org/soap/envelope/"
    $tempUriNs = "http://tempuri.org/wsGenRateEstimate/"

    $xmlWriter.WriteStartElement("soap", "Envelope", $soapEnvNs) 
    $xmlWriter.WriteAttributeString("xmlns", "xsi", $null, "http://www.w3.org/2001/XMLSchema-instance")
    $xmlWriter.WriteAttributeString("xmlns", "xsd", $null, "http://www.w3.org/2001/XMLSchema")

    $xmlWriter.WriteStartElement("soap", "Body", $soapEnvNs)
    $xmlWriter.WriteStartElement("RateEstimateRequestVO", $tempUriNs) 

    $xmlWriter.WriteElementString("Token", $tempUriNs, (Escape-Xml $apiToken))
    if (-not [string]::IsNullOrWhiteSpace($customerNumber)) {
        $xmlWriter.WriteElementString("CustomerNumber", $tempUriNs, (Escape-Xml $customerNumber))
    }
    $xmlWriter.WriteElementString("OriginCity", $tempUriNs, (Escape-Xml $ShipmentData.OriginCity))
    $xmlWriter.WriteElementString("OriginState", $tempUriNs, (Escape-Xml $ShipmentData.OriginState))
    $xmlWriter.WriteElementString("OriginZip", $tempUriNs, (Escape-Xml $ShipmentData.OriginZip))
    $xmlWriter.WriteElementString("OriginCountryCode", $tempUriNs, (Escape-Xml $ShipmentData.OriginCountryCode))
    $xmlWriter.WriteElementString("DestinationCity", $tempUriNs, (Escape-Xml $ShipmentData.DestinationCity))
    $xmlWriter.WriteElementString("DestinationState", $tempUriNs, (Escape-Xml $ShipmentData.DestinationState))
    $xmlWriter.WriteElementString("DestinationZip", $tempUriNs, (Escape-Xml $ShipmentData.DestinationZip))
    $xmlWriter.WriteElementString("DestinCountryCode", $tempUriNs, (Escape-Xml $ShipmentData.DestinCountryCode))
    $xmlWriter.WriteElementString("WhoAmI", $tempUriNs, (Escape-Xml $apiWhoAmI)) 
    $xmlWriter.WriteElementString("BillDate", $tempUriNs, (Escape-Xml $ShipmentData.BillDate))
        
    Write-Host "DEBUG PrePaidCollect VALUE CHECK: About to write '$($prePaidCollect)' to XML."
    $xmlWriter.WriteElementString("PrePaidCollect", $tempUriNs, (Escape-Xml $prePaidCollect))
 
    $totalPalletCountValue = if ($null -ne $ShipmentData.TotalPalletCountForPayload) { $ShipmentData.TotalPalletCountForPayload.ToString() } else { "0" } 
    $xmlWriter.WriteElementString("TotalPalletCount", $tempUriNs, (Escape-Xml $totalPalletCountValue))
    
    if ($ShipmentData.PrimaryAccCode -eq "PALET") {
        $xmlWriter.WriteStartElement("AccLine", $tempUriNs)
        $xmlWriter.WriteElementString("AccCode", $tempUriNs, "PALET")
        $xmlWriter.WriteEndElement() # AccLine
    }

    if ($ShipmentData.RateEstimateRequestLine -and $ShipmentData.RateEstimateRequestLine.Count -gt 0) {
        foreach ($lineItem in $ShipmentData.RateEstimateRequestLine) {
            # CRITICAL CHECK: Ensure Class is present if Weight is present for this line item
            if (($null -ne $lineItem.Weight -and $lineItem.Weight -ne "") -and ($null -eq $lineItem.Class -or $lineItem.Class -eq "" -or $lineItem.Class -match "^\s*$")) {
                Write-Warning "Invoke-AAACooperApi: FATAL - Attempting to send line item with Weight but no Class. Weight: '$($lineItem.Weight)', Class: '$($lineItem.Class)'. Skipping this API call for tariff '$tariffNameForLog'."
                return $null # Stop processing this API call if a line is invalid
            }

            $xmlWriter.WriteStartElement("RateEstimateRequestLine", $tempUriNs)
            
            $xmlWriter.WriteElementString("Weight", $tempUriNs, (Escape-Xml $lineItem.Weight))
            $xmlWriter.WriteElementString("Class", $tempUriNs, (Escape-Xml $lineItem.Class)) 
            
            $xmlWriter.WriteElementString("NMFC", $tempUriNs, (Escape-Xml $lineItem.NMFC))
            $xmlWriter.WriteElementString("NMFCSub", $tempUriNs, (Escape-Xml $lineItem.NMFCSub))

            $xmlWriter.WriteElementString("HandlingUnits", $tempUriNs, (Escape-Xml $lineItem.HandlingUnits)) 
            $xmlWriter.WriteElementString("HandlingUnitType", $tempUriNs, (Escape-Xml $lineItem.HandlingUnitType)) 
            $xmlWriter.WriteElementString("Hazmat", $tempUriNs, (Escape-Xml $lineItem.HazMat)) 
            
            $xmlWriter.WriteElementString("CubeU", $tempUriNs, (Escape-Xml $lineItem.CubeU)) 
            $xmlWriter.WriteElementString("Length", $tempUriNs, (Escape-Xml $lineItem.Length))  
            $xmlWriter.WriteElementString("Width", $tempUriNs, (Escape-Xml $lineItem.Width))   
            $xmlWriter.WriteElementString("Height", $tempUriNs, (Escape-Xml $lineItem.Height))  
            
            $xmlWriter.WriteEndElement() # RateEstimateRequestLine
        }
    } else { 
         $xmlWriter.WriteStartElement("RateEstimateRequestLine", $tempUriNs)
         $xmlWriter.WriteEndElement() # RateEstimateRequestLine
    }

    $xmlWriter.WriteEndElement() # RateEstimateRequestVO
    $xmlWriter.WriteEndElement() # soap:Body
    $xmlWriter.WriteEndElement() # soap:Envelope
    
    $xmlWriter.Flush(); $xmlWriter.Close()

    $soapRequestXmlString = '<?xml version="1.0" encoding="utf-8"?>' + $stringBuilder.ToString()
    
    Write-Verbose "AAA Cooper SOAP Request XML for Tariff '$tariffNameForLog':`n$soapRequestXmlString" 

    $apiUrl = $script:aaaCooperApiUri 
    if ([string]::IsNullOrWhiteSpace($apiUrl)) { Write-Error "AAA Cooper API URI not configured."; return $null }
    
    $utf8Bytes = [System.Text.Encoding]::UTF8.GetBytes($soapRequestXmlString)
    $headers = @{ "Content-Type" = "text/xml;charset=UTF-8" }

    try {
        $ProgressPreference = 'SilentlyContinue'
        $response = Invoke-WebRequest -Uri $apiUrl -Method Post -Headers $headers -Body $utf8Bytes -UseBasicParsing -ErrorAction Stop
        
        $responseStream = $response.RawContentStream
        $responseEncoding = if ($response.CharacterSet) { try {[System.Text.Encoding]::GetEncoding($response.CharacterSet)} catch {[System.Text.Encoding]::UTF8} } else {[System.Text.Encoding]::UTF8}
        $streamReader = New-Object System.IO.StreamReader($responseStream, $responseEncoding)
        $responseContent = $streamReader.ReadToEnd()
        $streamReader.Close(); $responseStream.Close()

        [xml]$responseXml = $responseContent
        Write-Verbose "AAA Cooper Raw Response XML for Tariff '$tariffNameForLog':`n$($responseXml.OuterXml)"

        $nsManager = New-Object System.Xml.XmlNamespaceManager($responseXml.NameTable)
        $nsManager.AddNamespace("soap", "http://schemas.xmlsoap.org/soap/envelope/")
        $nsManager.AddNamespace("res", $tempUriNs) 

        $faultNode = $responseXml.SelectSingleNode("/soap:Envelope/soap:Body/soap:Fault", $nsManager)
        if ($faultNode) {
            $faultString = $faultNode.SelectSingleNode("faultstring", $nsManager).InnerText
            $faultCode = $faultNode.SelectSingleNode("faultcode", $nsManager).InnerText
            Write-Warning "AAA Cooper API returned SOAP Fault for Tariff '$tariffNameForLog': Code='$faultCode', String='$faultString'"
            return $null
        }

        $rateResponseNode = $responseXml.SelectSingleNode("/soap:Envelope/soap:Body/res:RateEstimateResponseVO", $nsManager)
        if (-not $rateResponseNode) { $rateResponseNode = $responseXml.SelectSingleNode("/soap:Envelope/soap:Body/*[local-name()='RateEstimateResponseVO']", $nsManager)}

        if (-not $rateResponseNode) {
            Write-Warning "AAA Cooper API Response Error (${tariffNameForLog}): RateEstimateResponseVO node not found." 
            Write-Warning "Full Response (if RateEstimateResponseVO missing): $($responseXml.OuterXml)" 
            return $null
        }

        $errorMessageNode = $rateResponseNode.SelectSingleNode("res:ErrorMessage", $nsManager)
        if (-not $errorMessageNode) { $errorMessageNode = $rateResponseNode.SelectSingleNode("*[local-name()='ErrorMessage']", $nsManager)}

        if ($errorMessageNode -and -not [string]::IsNullOrWhiteSpace($errorMessageNode.InnerText)) {
            Write-Warning "AAA Cooper API Error (${tariffNameForLog}): $($errorMessageNode.InnerText)" 
            $quoteNumberNode = $responseXml.SelectSingleNode("res:QuoteNumber", $nsManager) 
            if (-not $quoteNumberNode) {$quoteNumberNode = $rateResponseNode.SelectSingleNode("*[local-name()='QuoteNumber']", $nsManager)}
            if ($quoteNumberNode -and -not [string]::IsNullOrWhiteSpace($quoteNumberNode.InnerText)) {
                Write-Warning "  -> QuoteNumber (with error): $($quoteNumberNode.InnerText)"
            }
            return $null 
        }

        $totalChargesNode = $responseXml.SelectSingleNode("res:TotalCharges", $nsManager) 
        if (-not $totalChargesNode) { $totalChargesNode = $rateResponseNode.SelectSingleNode("*[local-name()='TotalCharges']", $nsManager)}

        if ($totalChargesNode -and -not [string]::IsNullOrWhiteSpace($totalChargesNode.InnerText)) {
            $totalChargeValue = $totalChargesNode.InnerText
            try {
                $cleanedRate = $totalChargeValue -replace '[$,]' 
                $decimalRate = [decimal]$cleanedRate
                if ($decimalRate -lt 0) { throw "Negative rate received from API." }
                Write-Verbose "AAA Cooper API Call OK ($tariffNameForLog): Rate = $decimalRate"
                return $decimalRate
            } catch {
                Write-Warning "AAA Cooper Rate Conversion Fail (${tariffNameForLog}): Cannot convert rate '$totalChargeValue' to decimal. Error: $($_.Exception.Message)" 
                return $null
            }
        } else {
            Write-Warning "AAA Cooper Rate Not Found in Response (${tariffNameForLog}): 'TotalCharges' field is missing, null, or empty in API response." 
            $quoteNumberNode = $responseXml.SelectSingleNode("res:QuoteNumber", $nsManager) 
            if (-not $quoteNumberNode) {$quoteNumberNode = $rateResponseNode.SelectSingleNode("*[local-name()='QuoteNumber']", $nsManager)}
            $daysInTransitNode = $responseXml.SelectSingleNode("res:DaysInTransit", $nsManager) 
            if (-not $daysInTransitNode) {$daysInTransitNode = $rateResponseNode.SelectSingleNode("*[local-name()='DaysInTransit']", $nsManager)}

            if ($quoteNumberNode -and -not [string]::IsNullOrWhiteSpace($quoteNumberNode.InnerText)) {
                Write-Warning "  -> QuoteNumber found (but no TotalCharges): $($quoteNumberNode.InnerText)"
            }
            if ($daysInTransitNode -and -not [string]::IsNullOrWhiteSpace($daysInTransitNode.InnerText)) {
                Write-Warning "  -> DaysInTransit found (but no TotalCharges): $($daysInTransitNode.InnerText)"
            }
            Write-Warning "  -> Review the full raw XML response (enable verbose logging by setting `$VerbosePreference = 'Continue'` or changing Write-Verbose to Write-Host for the raw XML log line) to understand why TotalCharges is missing."
            return $null
        }
    } catch {
        $errMsg = $_.Exception.Message
        $statusCode = "N/A"; $eBody = "N/A"
        if ($_.Exception.Response) {
            try { $statusCode = $_.Exception.Response.StatusCode.value__ } catch { }
            try {
                $responseStreamForError = $_.Exception.Response.GetResponseStream()
                $errorResponseEncoding = [System.Text.Encoding]::UTF8 
                if ($_.Exception.Response.CharacterSet) {
                    try { $errorResponseEncoding = [System.Text.Encoding]::GetEncoding($_.Exception.Response.CharacterSet) } catch { Write-Warning "Could not get encoding from error response CharacterSet '$($_.Exception.Response.CharacterSet)'. Defaulting to UTF8."}
                }
                $errorStreamReader = New-Object System.IO.StreamReader($responseStreamForError, $errorResponseEncoding)
                $eBody = $errorStreamReader.ReadToEnd()
                $errorStreamReader.Close(); $responseStreamForError.Close()
            } catch { $eBody = "(Error reading response body: $($_.Exception.Message))" }
        }
        $truncatedBody = if ($eBody.Length -gt 500) { $eBody.Substring(0, 500) + "..." } else { $eBody }
        $fullErrMsg = "AAA Cooper API Invoke-WebRequest FAILED (${tariffNameForLog}): Error: $errMsg (HTTP Status: $statusCode) Response Body: $truncatedBody" 
        Write-Warning $fullErrMsg
        return $null
    }
}

Write-Verbose "TMS AAA Cooper Helper Functions loaded."
