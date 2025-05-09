# TMS_Helpers_RL.ps1
# Description: Contains helper functions specific to R+L Carriers operations,
#              including data normalization and API interaction. Added debug logging and refined object handling.
#              This file should be dot-sourced by the main script(s) after TMS_Config.ps1.

# Assumes config variables like $script:rlApiUri are available from TMS_Config.ps1
# Assumes general helper functions (if any were used by these) are available.

# --- Helper function to escape XML special characters ---
# MOVED TO SCRIPT SCOPE for reliable access
function Escape-Xml ($string) {
    if ($null -eq $string) { return '' };
    return [System.Security.SecurityElement]::Escape($string.ToString()) # Ensure it's a string before escaping
}

# --- Data Normalization Functions ---

function Load-And-Normalize-RLData {
    param(
        [Parameter(Mandatory=$true)][string]$CsvPath
    )
    Write-Host "`nLoading R+L data: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan
    # Define required columns based on Invoke-RLApi needs (including single commodity mapping)
    $reqCols = @( "Origin Postal Code", "Destination Postal Code", "Total Weight", "Freight Class 1", "Total Units", "Total Length", "Total Width", "Total Height") # Added Total Units

    try {
        if (-not (Test-Path -Path $CsvPath -PathType Leaf)) { Write-Error "CSV file not found at '$CsvPath'."; return $null }

        $rawData = Import-Csv -Path $CsvPath -Delimiter ',' -ErrorAction Stop # Assuming comma delimiter
        Write-Host " -> Rows read from CSV: $($rawData.Count)." -ForegroundColor Gray
        if ($rawData.Count -eq 0) { Write-Warning "CSV empty."; return @() } # Return empty array for no data

        $headers = $rawData[0].PSObject.Properties.Name
        $missingReq = $reqCols | Where-Object { $_ -notin $headers }
        if ($missingReq.Count -gt 0) { Write-Error "CSV missing required R+L columns: $($missingReq -join ', ')"; return $null }

        $normData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing R+L data..." -ForegroundColor Gray
        $invalid = 0; $csvRowNum = 1 # For user-friendly messages, header is row 1
        foreach ($row in $rawData) {
            $currentDataRowForMessage = $csvRowNum + 1 # Actual data row number

            $oZipRaw=$row."Origin Postal Code"; $dZipRaw=$row."Destination Postal Code"; $wtStrRaw=$row."Total Weight";
            $clStrRaw = $null
            if ($row.PSObject.Properties['Freight Class 1'] -ne $null) {
                $clStrRaw = $row."Freight Class 1"
                Write-Verbose "DEBUG (Load-And-Normalize-RLData): DataRow ${currentDataRowForMessage} - Raw 'Freight Class 1' value = [$clStrRaw]"
            } else {
                 Write-Verbose "DEBUG (Load-And-Normalize-RLData): DataRow ${currentDataRowForMessage} - Column 'Freight Class 1' not found or is null."
            }
            $pcsStrRaw=$row."Total Units"; $lenStrRaw=$row."Total Length"; $widStrRaw=$row."Total Width"; $hgtStrRaw=$row."Total Height"

            $oZip=$oZipRaw.Trim(); $dZip=$dZipRaw.Trim(); $wtStr=$wtStrRaw.Trim();
            $clStr = if ($clStrRaw -ne $null) { $clStrRaw.Trim() } else { $clStrRaw }
            Write-Verbose "DEBUG (Load-And-Normalize-RLData): DataRow ${currentDataRowForMessage} - Trimmed 'Freight Class 1' value for \$clStr = [$clStr]"

            $pcsStr=$pcsStrRaw.Trim(); $lenStr=$lenStrRaw.Trim(); $widStr=$widStrRaw.Trim(); $hgtStr=$hgtStrRaw.Trim()
            $wtNum=$null; $pcsNum=$null; $lenNum=$null; $widNum=$null; $hgtNum=$null; $skipRow = $false

            # Basic Validation
            if ([string]::IsNullOrWhiteSpace($oZip) -or $oZip.Length -lt 5) { $invalid++; Write-Verbose "Skip RL DataRow ${currentDataRowForMessage}: Bad Origin Zip '$($oZipRaw)'"; $skipRow = $true }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($dZip) -or $dZip.Length -lt 5)) { $invalid++; Write-Verbose "Skip RL DataRow ${currentDataRowForMessage}: Bad Dest Zip '$($dZipRaw)'"; $skipRow = $true }
            # Class is validated in Invoke-RLApi
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($pcsStr)) { $invalid++; Write-Verbose "Skip RL DataRow ${currentDataRowForMessage}: Bad Pieces (Total Units) '$($pcsStrRaw)'"; $skipRow = $true }


            # Numeric Validation
            if (-not $skipRow) { try { $wtNum = [decimal]$wtStr; if($wtNum -le 0){throw "Weight must be positive."} } catch { $invalid++; Write-Verbose "Skip RL DataRow ${currentDataRowForMessage}: Bad Weight '$($wtStrRaw)' Error: $($_.Exception.Message)"; $skipRow = $true } }
            if (-not $skipRow) { try { $pcsNum = [int]$pcsStr; if($pcsNum -le 0){throw "Pieces must be positive."} } catch { $invalid++; Write-Verbose "Skip RL DataRow ${currentDataRowForMessage}: Bad Pieces (Total Units) '$($pcsStrRaw)' Error: $($_.Exception.Message)"; $skipRow = $true } }
            if (-not $skipRow) { try { $lenNum = [decimal]$lenStr; if($lenNum -le 0){throw "Length must be positive."} } catch { $invalid++; Write-Verbose "Skip RL DataRow ${currentDataRowForMessage}: Bad Length '$($lenStrRaw)' Error: $($_.Exception.Message)"; $skipRow = $true } }
            if (-not $skipRow) { try { $widNum = [decimal]$widStr; if($widNum -le 0){throw "Width must be positive."} } catch { $invalid++; Write-Verbose "Skip RL DataRow ${currentDataRowForMessage}: Bad Width '$($widStrRaw)' Error: $($_.Exception.Message)"; $skipRow = $true } }
            if (-not $skipRow) { try { $hgtNum = [decimal]$hgtStr; if($hgtNum -le 0){throw "Height must be positive."} } catch { $invalid++; Write-Verbose "Skip RL DataRow ${currentDataRowForMessage}: Bad Height '$($hgtStrRaw)' Error: $($_.Exception.Message)"; $skipRow = $true } }


            if ($skipRow) { $csvRowNum++; continue }

            # Create a single commodity item for reports, explicitly as PSCustomObject
            $commodityItem = [PSCustomObject]@{
                Class = $clStr      # R+L API uses 'Class'
                Weight = $wtNum     # R+L API uses 'Weight' (will be converted to float later)
                Length = $lenNum    # R+L API uses 'Length'
                Width = $widNum     # R+L API uses 'Width'
                Height = $hgtNum    # R+L API uses 'Height'
                Pieces = $pcsNum    # R+L API might need pieces, though not explicitly in basic item structure for XML
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
                CustomerData = if ($headers -contains 'CustomerData') { $row.CustomerData.Trim() } else { $null } # For R+L API
                QuoteType = if ($headers -contains 'QuoteType') { $row.QuoteType.Trim() } else { 'Domestic' }
                CODAmount = if ($headers -contains 'CODAmount') { try {[decimal]$row.CODAmount.Trim()} catch {$null} } else { 0.0 }
                OriginCountryCode = if ($headers -contains 'OriginCountryCode') { $row.OriginCountryCode.Trim() } else { 'USA' }
                DestinationCountryCode = if ($headers -contains 'DestinationCountryCode') { $row.DestinationCountryCode.Trim() } else { 'USA' }
                DeclaredValue = if ($headers -contains 'DeclaredValue') { try {[decimal]$row.DeclaredValue.Trim()} catch {$null} } else { 0.0 }
                # Also store the original CSV values if needed for other reports or direct display
                'Total Weight' = $wtNum
                'Freight Class 1' = $clStr
                'Total Units' = $pcsNum
            }
            $normData.Add($normalizedEntry)
            $csvRowNum++
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
        [Parameter(Mandatory=$true)] [array]$Commodities, # Array of PSCustomObjects
        [Parameter(Mandatory=$false)] [PSCustomObject]$ShipmentDetails = $null # For non-commodity details
    )

    $tariffNameForLog = if ($KeyData.ContainsKey('TariffFileName')) { $KeyData.TariffFileName } elseif ($KeyData.ContainsKey('Name')) { $KeyData.Name } else { "UnknownRLTariff" }

    $apiKeyToUse = $null; $customerDataToUse = $null;
    if ($KeyData.ContainsKey('APIKey')) { $apiKeyToUse = $KeyData.APIKey }
    if ($KeyData.ContainsKey('CustomerData')) { $customerDataToUse = $KeyData.CustomerData }

    # --- Validate Base Inputs ---
    $missingFields = [System.Collections.Generic.List[string]]::new()
    if ([string]::IsNullOrWhiteSpace($OriginZip)) { $missingFields.Add("OriginZip") }
    if ([string]::IsNullOrWhiteSpace($DestinationZip)) { $missingFields.Add("DestinationZip") }
    if ($null -eq $Commodities -or $Commodities.Count -eq 0) { $missingFields.Add("Commodities (array empty or null)") }
    if ([string]::IsNullOrWhiteSpace($apiKeyToUse)) { $missingFields.Add("APIKey (from KeyData)") }

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
        # Allow CustomerData from ShipmentDetails to override KeyData if present
        if ($ShipmentDetails.PSObject.Properties.Match('CustomerData') -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.CustomerData)) { $customerDataToUse = $ShipmentDetails.CustomerData }
    }

    # Validate required non-commodity fields
    if ([string]::IsNullOrWhiteSpace($OriginCityToUse) -or $OriginCityToUse -eq "UNKNOWN") {$missingFields.Add("OriginCity")}
    if ([string]::IsNullOrWhiteSpace($OriginStateToUse) -or $OriginStateToUse -eq "XX") {$missingFields.Add("OriginState")}
    if ([string]::IsNullOrWhiteSpace($DestinationCityToUse) -or $DestinationCityToUse -eq "UNKNOWN") {$missingFields.Add("DestinationCity")}
    if ([string]::IsNullOrWhiteSpace($DestinationStateToUse) -or $DestinationStateToUse -eq "XX") {$missingFields.Add("DestinationState")}

    # --- Validate Commodities and Build XML Snippet ---
    $itemsXmlSnippet = ""
    if ($Commodities -is [array] -and $Commodities.Count -gt 0) {
        $validItemCount = 0
        for($c = 0; $c -lt $Commodities.Count; $c++){
            $item = $Commodities[$c]
            $itemClass = $null; $itemWeight = $null; $itemLength = 1.0; $itemWidth = 1.0; $itemHeight = 1.0; # Defaults for optional dims
            $isValidItem = $true; $currentItemErrors = [System.Collections.Generic.List[string]]::new()

            # Check item type
            if ($item -is [System.Collections.IDictionary] -or $item -is [psobject]) {
                try {
                    # Class
                    $classValue = $null
                    if ($item.PSObject.Properties.Match('Class').Count -gt 0) { # Check if 'Class' property exists
                        $classValue = $item.Class
                        Write-Verbose "DEBUG (Invoke-RLApi): Item $($c+1) - Read 'Class' value = [$classValue]"
                    } else {
                        Write-Verbose "DEBUG (Invoke-RLApi): Item $($c+1) - Property 'Class' not found."
                        $isValidItem = $false; $currentItemErrors.Add("Missing 'Class' property")
                    }
                    if ($isValidItem -and (-not([string]::IsNullOrWhiteSpace($classValue)) -and $classValue -match '^\d+(\.\d+)?$')) {
                        $itemClass = $classValue
                    } elseif ($isValidItem) { # Only add error if property was found but value is bad
                        $isValidItem = $false; $currentItemErrors.Add("Invalid Class value ('$($classValue)')")
                    }

                    # Weight
                    if ($item.PSObject.Properties.Match('Weight').Count -gt 0 -and -not([string]::IsNullOrWhiteSpace($item.Weight)) -and ($item.Weight -as [decimal]) -ne $null -and [decimal]$item.Weight -gt 0) {
                        $itemWeight = [decimal]$item.Weight
                    } else { $isValidItem = $false; $currentItemErrors.Add("Invalid or Missing Weight ('$($item.Weight)')") }

                    # Optional Dims - use default if missing or invalid
                    if ($item.PSObject.Properties.Match('Length').Count -gt 0 -and -not([string]::IsNullOrWhiteSpace($item.Length)) -and ($item.Length -as [decimal]) -ne $null -and [decimal]$item.Length -gt 0) { $itemLength = [decimal]$item.Length }
                    if ($item.PSObject.Properties.Match('Width').Count -gt 0 -and -not([string]::IsNullOrWhiteSpace($item.Width)) -and ($item.Width -as [decimal]) -ne $null -and [decimal]$item.Width -gt 0) { $itemWidth = [decimal]$item.Width }
                    if ($item.PSObject.Properties.Match('Height').Count -gt 0 -and -not([string]::IsNullOrWhiteSpace($item.Height)) -and ($item.Height -as [decimal]) -ne $null -and [decimal]$item.Height -gt 0) { $itemHeight = [decimal]$item.Height }

                } catch {
                    $isValidItem = $false; $currentItemErrors.Add("Unexpected error accessing commodity properties: $($_.Exception.Message)")
                }

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
                } else {
                     $missingFields.Add("Item $($c+1): $($currentItemErrors -join ', ')")
                }
            } else {
                $missingFields.Add("Commodity item $($c+1) is not a valid object type (Type: $($item.GetType().FullName)).")
            }
        } # End foreach item
        if ($validItemCount -eq 0 -and $Commodities.Count -gt 0) { $missingFields.Add("No valid commodity items found after validation.") }
    }
    # End Commodity Validation

    if ($missingFields.Count -gt 0) {
        Write-Warning "RL Skip: Tariff '$tariffNameForLog' - Missing/Invalid required data: $($missingFields -join '; ')."
        return $null
    }

    # --- Construct SOAP Payload ---
    $soapEndpoint = $script:rlApiUri # Assumes $script:rlApiUri is loaded from TMS_Config.ps1
    if ([string]::IsNullOrWhiteSpace($soapEndpoint)) { Write-Error "R+L API URI ('$($soapEndpoint)') is not defined or empty."; return $null } # Changed to Write-Error
    $soapAction = "http://www.rlcarriers.com/GetRateQuote"
    $tnsNamespace = "http://www.rlcarriers.com/"
    $soapNamespace = "http://schemas.xmlsoap.org/soap/envelope/"

    # Escape-Xml function is now at script scope

    $soapRequestBody = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="$soapNamespace">
 <soap:Body>
   <tns:GetRateQuote xmlns:tns="$tnsNamespace">
     <tns:APIKey>$(Escape-Xml $apiKeyToUse)</tns:APIKey>
     <tns:request>
       <tns:CustomerData>$(Escape-Xml $customerDataToUse)</tns:CustomerData>
       <tns:QuoteType>$(Escape-Xml $QuoteTypeToUse)</tns:QuoteType>
       <tns:CODAmount>$($CODAmountToUse.ToString("F2"))</tns:CODAmount>
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
       <tns:DeclaredValue>$($DeclaredValueToUse.ToString("F2"))</tns:DeclaredValue>
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
                    $errorMsgNode = $quoteDetails.SelectSingleNode("rl:Errors/rl:string", $nsManager) # R+L returns errors in an array of strings
                    if($errorMsgNode){
                        $errorMessages = $quoteDetails.SelectNodes("rl:Errors/rl:string", $nsManager) | ForEach-Object {$_.InnerText}
                        Write-Warning "RL API Error in Response for Tariff $($tariffNameForLog): $($errorMessages -join '; ')"
                    }
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