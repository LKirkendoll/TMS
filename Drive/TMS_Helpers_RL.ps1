# TMS_Helpers_RL.ps1
# Description: Contains helper functions specific to R+L Carriers operations,
#              including data normalization and API interaction.
#              This file should be dot-sourced by the main script(s) after TMS_Config.ps1.

# Assumes config variables like $script:rlApiUri are available from TMS_Config.ps1
# Assumes general helper functions (if any were used by these) are available.

# --- Data Normalization Functions ---

function Load-And-Normalize-RLData {
    # NOTE: This function might need adjustments if your CSV structure changes
    # to support multiple commodities per row for R+L reports.
    # Currently, it assumes single commodity details per row based on original design.
    param(
        [Parameter(Mandatory=$true)][string]$CsvPath
    )
    Write-Host "`nLoading R+L data: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan
    # Define required columns based on Invoke-RLApi needs (including single commodity mapping)
    $reqCols = @( "Origin Postal Code", "Destination Postal Code", "Total Weight", "Freight Class 1", "Origin City", "Origin State", "Destination City", "Destination State", "Total Length", "Total Width", "Total Height")

    try {
        if (-not (Test-Path -Path $CsvPath -PathType Leaf)) { Write-Error "CSV file not found at '$CsvPath'."; return $null }
        $rawData = Import-Csv -Path $CsvPath -ErrorAction Stop
        Write-Host " -> Rows read from CSV: $($rawData.Count)." -ForegroundColor Gray
        if ($rawData.Count -eq 0) { Write-Warning "CSV empty."; return @() } # Return empty array for no data

        $headers = $rawData[0].PSObject.Properties.Name
        $missingReq = $reqCols | Where-Object { $_ -notin $headers }
        if ($missingReq.Count -gt 0) { Write-Error "CSV missing required R+L columns: $($missingReq -join ', ')"; return $null }

        $normData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing R+L data..." -ForegroundColor Gray
        $invalid = 0; $rowNum = 1
        foreach ($row in $rawData) {
            $rowNum++
            $oZipRaw=$row."Origin Postal Code"; $dZipRaw=$row."Destination Postal Code"; $wtStrRaw=$row."Total Weight"; $clStrRaw=$row."Freight Class 1"; $lenStrRaw=$row."Total Length"; $widStrRaw=$row."Total Width"; $hgtStrRaw=$row."Total Height"
            $oZip=$oZipRaw.Trim(); $dZip=$dZipRaw.Trim(); $wtStr=$wtStrRaw.Trim(); $clStr=$clStrRaw.Trim(); $lenStr=$lenStrRaw.Trim(); $widStr=$widStrRaw.Trim(); $hgtStr=$hgtStrRaw.Trim()
            $wtNum=$null; $lenNum=$null; $widNum=$null; $hgtNum=$null; $skipRow = $false

            # Basic Validation
            if ([string]::IsNullOrWhiteSpace($oZip) -or $oZip.Length -lt 5) { $invalid++; Write-Verbose "Skip RL Row ${rowNum}: Bad Origin Zip"; $skipRow = $true }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($dZip) -or $dZip.Length -lt 5)) { $invalid++; Write-Verbose "Skip RL Row ${rowNum}: Bad Dest Zip"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($clStr)) { $invalid++; Write-Verbose "Skip RL Row ${rowNum}: Bad Class"; $skipRow = $true }
            # Numeric Validation
            if (-not $skipRow) { try { $wtNum = [decimal]$wtStr; if($wtNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip RL Row ${rowNum}: Bad Weight"; $skipRow = $true } }
            if (-not $skipRow) { try { $lenNum = [decimal]$lenStr; if($lenNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip RL Row ${rowNum}: Bad Length"; $skipRow = $true } }
            if (-not $skipRow) { try { $widNum = [decimal]$widStr; if($widNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip RL Row ${rowNum}: Bad Width"; $skipRow = $true } }
            if (-not $skipRow) { try { $hgtNum = [decimal]$hgtStr; if($hgtNum -le 0){throw} } catch { $invalid++; Write-Verbose "Skip RL Row ${rowNum}: Bad Height"; $skipRow = $true } }


            if ($skipRow) { continue }

            # Create a single commodity item for reports
            $commodityItem = [ordered]@{
                Class = $clStr      # R+L API uses 'Class'
                Weight = $wtNum     # R+L API uses 'Weight' (will be converted to float later)
                Length = $lenNum    # R+L API uses 'Length'
                Width = $widNum     # R+L API uses 'Width'
                Height = $hgtNum    # R+L API uses 'Height'
                # Pieces not explicitly in R+L basic item structure, handled by weight/dims per item
            }

            # Create normalized object, grabbing optional fields if they exist
            $normalizedEntry = [PSCustomObject]@{
                OriginZip = $oZip
                DestinationZip = $dZip
                # Store commodities as an array
                Commodities = @($commodityItem)
                # Add other optional fields that Invoke-RLApi might use from the $ShipmentDetails parameter
                OriginCity = if ($headers -contains 'Origin City') { $row.'Origin City'.Trim() } else { $null }
                OriginState = if ($headers -contains 'Origin State') { $row.'Origin State'.Trim() } else { $null }
                DestinationCity = if ($headers -contains 'Destination City') { $row.'Destination City'.Trim() } else { $null }
                DestinationState = if ($headers -contains 'Destination State') { $row.'Destination State'.Trim() } else { $null }
                CustomerData = if ($headers -contains 'CustomerData') { $row.CustomerData.Trim() } else { $null }
                QuoteType = if ($headers -contains 'QuoteType') { $row.QuoteType.Trim() } else { 'Domestic' }
                CODAmount = if ($headers -contains 'CODAmount') { try {[decimal]$row.CODAmount.Trim()} catch {$null} } else { 0.0 }
                OriginCountryCode = if ($headers -contains 'OriginCountryCode') { $row.OriginCountryCode.Trim() } else { 'USA' }
                DestinationCountryCode = if ($headers -contains 'DestinationCountryCode') { $row.DestinationCountryCode.Trim() } else { 'USA' }
                DeclaredValue = if ($headers -contains 'DeclaredValue') { try {[decimal]$row.DeclaredValue.Trim()} catch {$null} } else { 0.0 }
            }
            $normData.Add($normalizedEntry)
        }
        if ($invalid -gt 0) { Write-Warning " -> Skipped $invalid R+L rows (missing/invalid data)." }
        Write-Host " -> OK: $($normData.Count) R+L rows normalized." -ForegroundColor Green; return $normData
    } catch { Write-Error "Error processing R+L CSV '$CsvPath': $($_.Exception.Message)"; return $null }
}

# --- API Call Functions ---

function Invoke-RLApi {
    param(
        [Parameter(Mandatory=$true)] [hashtable]$KeyData,
        [Parameter(Mandatory=$true)] [string]$OriginZip,
        [Parameter(Mandatory=$true)] [string]$DestinationZip,
        # <<< PARAMETER CHANGE: Accept array of commodities >>>
        [Parameter(Mandatory=$true)] [array]$Commodities, # Array of hashtables/PSObjects
        # Removed single Weight/Class params
        [Parameter(Mandatory=$false)] [PSCustomObject]$ShipmentDetails = $null # For non-commodity details like City, State, DeclaredValue etc.
    )

    $tariffNameForLog = if ($KeyData.ContainsKey('TariffFileName')) { $KeyData.TariffFileName } elseif ($KeyData.ContainsKey('Name')) { $KeyData.Name } else { "UnknownRLTariff" }

    $apiKeyToUse = $null; $customerDataToUse = $null;
    if ($KeyData.ContainsKey('APIKey')) { $apiKeyToUse = $KeyData.APIKey }
    if ($KeyData.ContainsKey('CustomerData')) { $customerDataToUse = $KeyData.CustomerData }

    # --- Validate Base Inputs ---
    $missingFields = @()
    if ([string]::IsNullOrWhiteSpace($OriginZip)) { $missingFields += "OriginZip" }
    if ([string]::IsNullOrWhiteSpace($DestinationZip)) { $missingFields += "DestinationZip" }
    if ($null -eq $Commodities -or $Commodities.Count -eq 0) { $missingFields += "Commodities (array empty or null)" }
    if ([string]::IsNullOrWhiteSpace($apiKeyToUse)) { $missingFields += "APIKey (from KeyData)" }

    # --- Get Non-Commodity Details from $ShipmentDetails if provided ---
    $OriginCityToUse = "UNKNOWN"; $OriginStateToUse = "XX"; $OriginCountryCodeToUse = "USA";
    $DestinationCityToUse = "UNKNOWN"; $DestinationStateToUse = "XX"; $DestinationCountryCodeToUse = "USA";
    $QuoteTypeToUse = "Domestic"; $CODAmountToUse = 0.0; $DeclaredValueToUse = 0.0;

    if ($null -ne $ShipmentDetails) {
        if ($ShipmentDetails.PSObject.Properties.Match('OriginCity') -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.OriginCity)) { $OriginCityToUse = $ShipmentDetails.OriginCity }
        if ($ShipmentDetails.PSObject.Properties.Match('OriginState') -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.OriginState)) { $OriginStateToUse = $ShipmentDetails.OriginState }
        if ($ShipmentDetails.PSObject.Properties.Match('OriginCountryCode') -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.OriginCountryCode)) { $OriginCountryCodeToUse = $ShipmentDetails.OriginCountryCode }
        if ($ShipmentDetails.PSObject.Properties.Match('DestinationCity') -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.DestinationCity)) { $DestinationCityToUse = $ShipmentDetails.DestinationCity }
        if ($ShipmentDetails.PSObject.Properties.Match('DestinationState') -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.DestinationState)) { $DestinationStateToUse = $ShipmentDetails.DestinationState }
        if ($ShipmentDetails.PSObject.Properties.Match('DestinationCountryCode') -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.DestinationCountryCode)) { $DestinationCountryCodeToUse = $ShipmentDetails.DestinationCountryCode }
        if ($ShipmentDetails.PSObject.Properties.Match('QuoteType') -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.QuoteType)) { $QuoteTypeToUse = $ShipmentDetails.QuoteType }
        if ($ShipmentDetails.PSObject.Properties.Match('CODAmount') -and $null -ne $ShipmentDetails.CODAmount) { try { $CODAmountToUse = [decimal]$ShipmentDetails.CODAmount } catch {} }
        if ($ShipmentDetails.PSObject.Properties.Match('DeclaredValue') -and $null -ne $ShipmentDetails.DeclaredValue) { try { $DeclaredValueToUse = [decimal]$ShipmentDetails.DeclaredValue } catch {} }
        if ($ShipmentDetails.PSObject.Properties.Match('CustomerData') -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.CustomerData)) { $customerDataToUse = $ShipmentDetails.CustomerData } # Allow override from shipment details
    }

    # Validate required non-commodity fields
    if ([string]::IsNullOrWhiteSpace($OriginCityToUse) -or $OriginCityToUse -eq "UNKNOWN") {$missingFields += "OriginCity"}
    if ([string]::IsNullOrWhiteSpace($OriginStateToUse) -or $OriginStateToUse -eq "XX") {$missingFields += "OriginState"}
    if ([string]::IsNullOrWhiteSpace($DestinationCityToUse) -or $DestinationCityToUse -eq "UNKNOWN") {$missingFields += "DestinationCity"}
    if ([string]::IsNullOrWhiteSpace($DestinationStateToUse) -or $DestinationStateToUse -eq "XX") {$missingFields += "DestinationState"}

    # --- Validate Commodities and Build XML Snippet ---
    $itemsXmlSnippet = ""
    if ($Commodities -is [array] -and $Commodities.Count -gt 0) {
        $validItemCount = 0
        foreach ($item in $Commodities) {
            if ($item -isnot [hashtable] -and $item -isnot [psobject]) { $missingFields += "Commodity item not valid object."; continue }

            $itemClass = $null; $itemWeight = $null; $itemLength = 1.0; $itemWidth = 1.0; $itemHeight = 1.0; # Defaults for optional dims
            $isValidItem = $true

            if ($item.PSObject.Properties.Match('Class') -and -not [string]::IsNullOrWhiteSpace($item.Class)) { $itemClass = $item.Class } else { $isValidItem = $false; $missingFields += "Item missing/invalid Class '$($item.Class)'" }
            if ($item.PSObject.Properties.Match('Weight') -and $item.Weight -as [decimal] -ne $null -and [decimal]$item.Weight -gt 0) { $itemWeight = [decimal]$item.Weight } else { $isValidItem = $false; $missingFields += "Item missing/invalid Weight '$($item.Weight)'" }
            # Optional Dims - use default if missing or invalid
            if ($item.PSObject.Properties.Match('Length') -and $item.Length -as [decimal] -ne $null -and [decimal]$item.Length -gt 0) { $itemLength = [decimal]$item.Length }
            if ($item.PSObject.Properties.Match('Width') -and $item.Width -as [decimal] -ne $null -and [decimal]$item.Width -gt 0) { $itemWidth = [decimal]$item.Width }
            if ($item.PSObject.Properties.Match('Height') -and $item.Height -as [decimal] -ne $null -and [decimal]$item.Height -gt 0) { $itemHeight = [decimal]$item.Height }

            if ($isValidItem) {
                $validItemCount++
                # Convert numbers to appropriate string format for XML (float for R+L)
                $weightStr = ([float]$itemWeight).ToString("F2") # Example: 2 decimal places
                $lengthStr = ([float]$itemLength).ToString("F2")
                $widthStr = ([float]$itemWidth).ToString("F2")
                $heightStr = ([float]$itemHeight).ToString("F2")

                # Append XML for this item
                $itemsXmlSnippet += @"
         <tns:Item>
           <tns:Class>$(Escape-Xml $itemClass)</tns:Class>
           <tns:Weight>$weightStr</tns:Weight>
           <tns:Width>$widthStr</tns:Width>
           <tns:Height>$heightStr</tns:Height>
           <tns:Length>$lengthStr</tns:Length>
         </tns:Item>
"@
            }
        } # End foreach item
        if ($validItemCount -eq 0) { $missingFields += "No valid commodity items found after validation." }
    }
    # End Commodity Validation

    if ($missingFields.Count -gt 0) {
        Write-Warning "RL Skip: Tariff '$tariffNameForLog' - Missing/Invalid required data: $($missingFields -join ', ')."
        return $null
    }

    # --- Construct SOAP Payload ---
    $soapEndpoint = $script:rlApiUri
    if ([string]::IsNullOrWhiteSpace($soapEndpoint)) { throw "R+L API URI ('$($soapEndpoint)') is not defined or empty." }
    $soapAction = "http://www.rlcarriers.com/GetRateQuote"
    $tnsNamespace = "http://www.rlcarriers.com/"
    $soapNamespace = "http://schemas.xmlsoap.org/soap/envelope/"

    # Helper to escape XML special characters
    function Escape-Xml ($string) { if ($null -eq $string) { return '' }; return [System.Security.SecurityElement]::Escape($string) }

    $soapRequestBody = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="$soapNamespace">
 <soap:Body>
   <tns:GetRateQuote xmlns:tns="$tnsNamespace">
     <tns:APIKey>$(Escape-Xml $apiKeyToUse)</tns:APIKey>
     <tns:request>
       <tns:CustomerData>$(Escape-Xml $customerDataToUse)</tns:CustomerData>
       <tns:QuoteType>$(Escape-Xml $QuoteTypeToUse)</tns:QuoteType>
       <tns:CODAmount>$CODAmountToUse</tns:CODAmount>
       <tns:Origin>
         <tns:City>$(Escape-Xml $OriginCityToUse)</tns:City>
         <tns:StateOrProvince>$(Escape-Xml $OriginStateToUse)</tns:StateOrProvince>
         <tns:ZipOrPostalCode>$(Escape-Xml $OriginZip)</tns:ZipOrPostalCode>
         <tns:CountryCode>$(Escape-Xml $OriginCountryCodeToUse)</tns:CountryCode>
       </tns:Origin>
       <tns:Destination>
         <tns:City>$(Escape-Xml $DestinationCityToUse)</tns:City>
         <tns:StateOrProvince>$(Escape-Xml $DestinationStateToUse)</tns:StateOrProvince>
         <tns:ZipOrPostalCode>$(Escape-Xml $DestinationZip)</tns:ZipOrPostalCode>
         <tns:CountryCode>$(Escape-Xml $DestinationCountryCodeToUse)</tns:CountryCode>
       </tns:Destination>
       <tns:Items>$itemsXmlSnippet</tns:Items>
       <tns:DeclaredValue>$DeclaredValueToUse</tns:DeclaredValue>
       </tns:request>
   </tns:GetRateQuote>
 </soap:Body>
</soap:Envelope>
"@

    # --- API Call ---
    $headers = @{ "Content-Type" = "text/xml; charset=utf-8"; "SOAPAction" = "`"$soapAction`"" }
    Write-Verbose "Calling RL API: Tariff $tariffNameForLog"
    try {
        $ProgressPreference = 'SilentlyContinue' # Suppress progress bar from Invoke-WebRequest
        $response = Invoke-WebRequest -Uri $soapEndpoint -Method Post -Headers $headers -Body $soapRequestBody -UseBasicParsing -ErrorAction Stop
        [xml]$responseXml = $response.Content

        # --- Process Response ---
        $nsManager = New-Object System.Xml.XmlNamespaceManager($responseXml.NameTable)
        $nsManager.AddNamespace("soap", $soapNamespace); $nsManager.AddNamespace("rl", $tnsNamespace)

        # Check for SOAP Fault
        $faultNode = $responseXml.SelectSingleNode("/soap:Envelope/soap:Body/soap:Fault", $nsManager)
        if ($faultNode) {
            $faultString = $faultNode.SelectSingleNode("faultstring", $nsManager).InnerText
            $faultCode = $faultNode.SelectSingleNode("faultcode", $nsManager).InnerText
            Write-Warning "RL API returned SOAP Fault for Tariff $($tariffNameForLog): Code='$faultCode', String='$faultString'"
            Write-Verbose "RL Response XML (Fault - Tariff $tariffNameForLog): $($response.Content)"
            return $null
        }

        # Find the result node
        $rateQuoteResult = $responseXml.SelectSingleNode("/soap:Envelope/soap:Body/rl:GetRateQuoteResponse/rl:GetRateQuoteResult", $nsManager)
        if ($rateQuoteResult) {
            $quoteDetails = $rateQuoteResult.SelectSingleNode("rl:Result", $nsManager)
            if ($quoteDetails) {
                # Find the NET charge amount
                $netChargeEntry = $quoteDetails.SelectSingleNode("rl:Charges/rl:Charge[rl:Type='NET']/rl:Amount", $nsManager)
                if ($netChargeEntry) {
                    $totalChargeValue = $netChargeEntry.InnerText
                    try {
                        $cleanedRate = $totalChargeValue -replace '[$,]'
                        $decimalRate = [decimal]$cleanedRate
                        if ($decimalRate -lt 0) { throw "Negative rate returned."}
                        Write-Verbose "RL OK: Tariff $tariffNameForLog Rate: $decimalRate"
                        return $decimalRate
                    } catch { Write-Warning "RL Convert Fail for Tariff $($tariffNameForLog): Cannot convert rate '$totalChargeValue' to decimal. Error: $($_.Exception.Message)"; return $null }
                } else {
                    # Check for specific errors returned by R+L if NET charge missing
                    $errorMsgNode = $quoteDetails.SelectSingleNode("rl:Errors/rl:string", $nsManager)
                    if($errorMsgNode){ Write-Warning "RL API Error in Response for Tariff $($tariffNameForLog): $($errorMsgNode.InnerText)" }
                    else { Write-Warning "RL Resp Missing 'NET' charge for Tariff $tariffNameForLog. Check response XML." }
                    Write-Verbose "RL Response XML (No NET Charge - Tariff $tariffNameForLog): $($response.Content)"
                    return $null
                }
            } else { Write-Warning "RL Resp structure unexpected (No 'Result') for Tariff $tariffNameForLog."; Write-Verbose "RL Response XML: $($response.Content)"; return $null }
        } else { Write-Warning "RL Resp structure unexpected (No 'GetRateQuoteResult' or 'Fault') for Tariff $tariffNameForLog."; Write-Verbose "RL Response XML: $($response.Content)"; return $null }
    } catch {
        $errMsg = $_.Exception.Message; $statusCode = "N/A"; $eBody = "N/A"
        if ($_.Exception.Response) { try {$statusCode = $_.Exception.Response.StatusCode.value__} catch{}; try { $stream = $_.Exception.Response.GetResponseStream(); $reader = New-Object System.IO.StreamReader($stream); $eBody = $reader.ReadToEnd(); $reader.Close(); $stream.Close() } catch {$eBody="(Err reading resp body)"} }
        $truncatedBody = if ($eBody.Length -gt 500) { $eBody.Substring(0, 500) + "..." } else { $eBody }
        $fullErrMsg = "RL FAIL: Tariff $tariffNameForLog. Error: $errMsg (HTTP $statusCode) Resp: $truncatedBody"
        Write-Warning $fullErrMsg; return $null
    }
}

Write-Verbose "TMS R+L Helper Functions loaded."