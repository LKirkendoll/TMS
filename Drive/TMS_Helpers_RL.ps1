# TMS_Helpers_RL.ps1
# Description: Contains helper functions specific to R+L Carriers operations,
#              including data normalization and API interaction.
#              This file should be dot-sourced by the main script(s) after TMS_Config.ps1.

# Assumes config variables like $script:rlApiUri are available from TMS_Config.ps1
# Assumes general helper functions (if any were used by these) are available.

# --- Data Normalization Functions ---

function Load-And-Normalize-RLData {
    param(
        [Parameter(Mandatory=$true)][string]$CsvPath
    )
    Write-Host "`nLoading R+L data: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan
    # Define required columns based on Invoke-RLApi needs
    $reqCols = @( "Origin Postal Code", "Destination Postal Code", "Total Weight", "Freight Class 1", "Origin City", "Origin State", "Destination City", "Destination State") 

    try {
        if (-not (Test-Path -Path $CsvPath -PathType Leaf)) {
            Write-Error "CSV file not found at '$CsvPath'."
            return $null
        }
        $rawData = Import-Csv -Path $CsvPath -ErrorAction Stop
        Write-Host " -> Rows read from CSV: $($rawData.Count)." -ForegroundColor Gray
        if ($rawData.Count -eq 0) { Write-Warning "CSV empty."; return @() } # Return empty array for no data
        
        $headers = $rawData[0].PSObject.Properties.Name
        $missingReq = $reqCols | Where-Object { $_ -notin $headers }
        if ($missingReq.Count -gt 0) { Write-Error "CSV missing required R+L columns: $($missingReq -join ', ')"; return $null }
        
        # Optional columns check (if you define $optCols for R+L)
        # $optCols = @( ... ) 
        # $missingOpt = $optCols | Where-Object { $_ -notin $headers }; if($missingOpt.Count -gt 0){ Write-Warning "CSV missing optional R+L columns: $($missingOpt -join ', ')" }

        $normData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing R+L data..." -ForegroundColor Gray
        $invalid = 0; $rowNum = 1
        foreach ($row in $rawData) {
            $rowNum++
            $oZipRaw=$row."Origin Postal Code"; $dZipRaw=$row."Destination Postal Code"; $wtStrRaw=$row."Total Weight"; $clStrRaw=$row."Freight Class 1"
            $oZip=$oZipRaw.Trim(); $dZip=$dZipRaw.Trim(); $wtStr=$wtStrRaw.Trim(); $clStr=$clStrRaw.Trim()
            $wtNum=$null; $skipRow = $false

            if ([string]::IsNullOrWhiteSpace($oZip) -or $oZip.Length -lt 5) { $invalid++; Write-Verbose "Skip RL Row ${rowNum}: Bad Origin Zip '$oZipRaw'"; $skipRow = $true }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($dZip) -or $dZip.Length -lt 5)) { $invalid++; Write-Verbose "Skip RL Row ${rowNum}: Bad Dest Zip '$dZipRaw'"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($clStr)) { $invalid++; Write-Verbose "Skip RL Row ${rowNum}: Bad Class '$clStrRaw'"; $skipRow = $true }
            if (-not $skipRow) { 
                try { 
                    $wtNum = [decimal]$wtStr
                    if($wtNum -le 0){throw "Weight must be positive."} 
                } catch { 
                    $invalid++; Write-Verbose "Skip RL Row ${rowNum}: Bad Weight '$wtStrRaw' Error: $($_.Exception.Message)"; $skipRow = $true 
                } 
            }

            if ($skipRow) { continue }

            # Create normalized object, grabbing optional fields if they exist
            # Ensure consistency with how properties are accessed in Invoke-RLApi
            $normalizedEntry = [PSCustomObject]@{
                OriginZip = $oZip
                DestinationZip = $dZip
                Weight = $wtNum # Store numeric weight
                Class = $clStr
                OriginCity = if ($headers -contains 'Origin City') { $row.'Origin City'.Trim() } elseif ($headers -contains 'OriginCity') { $row.OriginCity.Trim() } else { $null }
                OriginState = if ($headers -contains 'Origin State') { $row.'Origin State'.Trim() } elseif ($headers -contains 'OriginState') { $row.OriginState.Trim() } else { $null }
                DestinationCity = if ($headers -contains 'Destination City') { $row.'Destination City'.Trim() } elseif ($headers -contains 'DestinationCity') { $row.DestinationCity.Trim() } else { $null }
                DestinationState = if ($headers -contains 'Destination State') { $row.'Destination State'.Trim() } elseif ($headers -contains 'DestinationState') { $row.DestinationState.Trim() } else { $null }
                # Add other optional fields that Invoke-RLApi might use from the $ShipmentDetails parameter
                CustomerData = if ($headers -contains 'CustomerData') { $row.CustomerData.Trim() } else { $null }
                QuoteType = if ($headers -contains 'QuoteType') { $row.QuoteType.Trim() } else { 'Domestic' } 
                CODAmount = if ($headers -contains 'CODAmount') { try {[decimal]$row.CODAmount.Trim()} catch {$null} } else { 0.0 }
                OriginCountryCode = if ($headers -contains 'OriginCountryCode') { $row.OriginCountryCode.Trim() } else { 'USA' } 
                DestinationCountryCode = if ($headers -contains 'DestinationCountryCode') { $row.DestinationCountryCode.Trim() } else { 'USA' } 
                ItemWidth = if ($headers -contains 'ItemWidth') { try {[float]$row.ItemWidth.Trim()} catch {1.0} } else { 1.0 } 
                ItemHeight = if ($headers -contains 'ItemHeight') { try {[float]$row.ItemHeight.Trim()} catch {1.0} } else { 1.0 }
                ItemLength = if ($headers -contains 'ItemLength') { try {[float]$row.ItemLength.Trim()} catch {1.0} } else { 1.0 }
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
        [Parameter(Mandatory=$true)] [decimal]$Weight, 
        [Parameter(Mandatory=$true)] [string]$Class,   
        [Parameter(Mandatory=$false)] [PSCustomObject]$ShipmentDetails = $null 
    )

    $tariffNameForLog = if ($KeyData.ContainsKey('TariffFileName')) { $KeyData.TariffFileName } else { "UnknownTariff" }
    if (-not $KeyData.ContainsKey('Name')) { $KeyData.Name = $tariffNameForLog } 

    $apiKeyToUse = $null; $customerDataToUse = $null;
    if ($KeyData.ContainsKey('APIKey')) { $apiKeyToUse = $KeyData.APIKey }
    if ($KeyData.ContainsKey('CustomerData')) { $customerDataToUse = $KeyData.CustomerData } 

    $missingFields = @()
    if ([string]::IsNullOrWhiteSpace($OriginZip)) { $missingFields += "OriginZip" }
    if ([string]::IsNullOrWhiteSpace($DestinationZip)) { $missingFields += "DestinationZip" }
    if ($null -eq $Weight -or $Weight -le 0) { $missingFields += "Weight(<=0 or invalid: '$Weight')" }
    if ([string]::IsNullOrWhiteSpace($Class)) { $missingFields += "Class" }
    if ([string]::IsNullOrWhiteSpace($apiKeyToUse)) { $missingFields += "APIKey (from KeyData)" }
    
    # R+L API requires City and State for Origin and Destination
    $OriginCityToUse = "UNKNOWN"; $OriginStateToUse = "XX"; $OriginCountryCodeToUse = "USA";
    $DestinationCityToUse = "UNKNOWN"; $DestinationStateToUse = "XX"; $DestinationCountryCodeToUse = "USA";
    $QuoteTypeToUse = "Domestic"; $CODAmountToUse = 0.0; $DeclaredValueToUse = 0.0;
    $ItemWidthToUse = 1.0; $ItemHeightToUse = 1.0; $ItemLengthToUse = 1.0;

    if ($null -ne $ShipmentDetails) {
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'OriginCity' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.OriginCity)) { $OriginCityToUse = $ShipmentDetails.OriginCity }
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'OriginState' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.OriginState)) { $OriginStateToUse = $ShipmentDetails.OriginState }
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'OriginCountryCode' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.OriginCountryCode)) { $OriginCountryCodeToUse = $ShipmentDetails.OriginCountryCode }
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'DestinationCity' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.DestinationCity)) { $DestinationCityToUse = $ShipmentDetails.DestinationCity }
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'DestinationState' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.DestinationState)) { $DestinationStateToUse = $ShipmentDetails.DestinationState }
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'DestinationCountryCode' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.DestinationCountryCode)) { $DestinationCountryCodeToUse = $ShipmentDetails.DestinationCountryCode }
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'QuoteType' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.QuoteType)) { $QuoteTypeToUse = $ShipmentDetails.QuoteType }
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'CODAmount' -and $null -ne $ShipmentDetails.CODAmount) { try { $CODAmountToUse = [decimal]$ShipmentDetails.CODAmount } catch {} }
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'DeclaredValue' -and $null -ne $ShipmentDetails.DeclaredValue) { try { $DeclaredValueToUse = [decimal]$ShipmentDetails.DeclaredValue } catch {} }
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'ItemWidth' -and $null -ne $ShipmentDetails.ItemWidth) { try { $ItemWidthToUse = [float]$ShipmentDetails.ItemWidth } catch { $ItemWidthToUse=1.0 } }
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'ItemHeight' -and $null -ne $ShipmentDetails.ItemHeight) { try { $ItemHeightToUse = [float]$ShipmentDetails.ItemHeight } catch { $ItemHeightToUse=1.0 } }
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'ItemLength' -and $null -ne $ShipmentDetails.ItemLength) { try { $ItemLengthToUse = [float]$ShipmentDetails.ItemLength } catch { $ItemLengthToUse=1.0 } }
        if ($ShipmentDetails.PSObject.Properties.Name -contains 'CustomerData' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.CustomerData)) { $customerDataToUse = $ShipmentDetails.CustomerData }
    } else {
         Write-Verbose "RL Invoke: No ShipmentDetails provided for optional fields (City/State etc.) for Tariff '$tariffNameForLog'. Using defaults."
    }
    
    if ([string]::IsNullOrWhiteSpace($OriginCityToUse) -or $OriginCityToUse -eq "UNKNOWN") {$missingFields += "OriginCity"}
    if ([string]::IsNullOrWhiteSpace($OriginStateToUse) -or $OriginStateToUse -eq "XX") {$missingFields += "OriginState"}
    if ([string]::IsNullOrWhiteSpace($DestinationCityToUse) -or $DestinationCityToUse -eq "UNKNOWN") {$missingFields += "DestinationCity"}
    if ([string]::IsNullOrWhiteSpace($DestinationStateToUse) -or $DestinationStateToUse -eq "XX") {$missingFields += "DestinationState"}

    if ($missingFields.Count -gt 0) {
        Write-Warning "RL Skip: Tariff '$tariffNameForLog' - Missing required data: $($missingFields -join ', ')."
        return $null
    }

    $ItemWeightToUse = try {[float]$Weight} catch { Write-Warning "Could not convert Weight '$Weight' to float for RL item payload."; 0.0 }
    if ($ItemWeightToUse -le 0) {
        Write-Warning "RL Skip: Invalid Item Weight ($ItemWeightToUse) for Tariff '$tariffNameForLog'."
        return $null
    }

    $soapEndpoint = $script:rlApiUri 
    if ([string]::IsNullOrWhiteSpace($soapEndpoint)) { throw "R+L API URI ('$($soapEndpoint)') is not defined or empty." }
    $soapAction = "http://www.rlcarriers.com/GetRateQuote"
    $tnsNamespace = "http://www.rlcarriers.com/"
    $soapNamespace = "http://schemas.xmlsoap.org/soap/envelope/"

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
       <tns:Items>
         <tns:Item>
           <tns:Class>$(Escape-Xml $Class)</tns:Class>
           <tns:Weight>$ItemWeightToUse</tns:Weight>
           <tns:Width>$ItemWidthToUse</tns:Width>
           <tns:Height>$ItemHeightToUse</tns:Height>
           <tns:Length>$ItemLengthToUse</tns:Length>
         </tns:Item>
         </tns:Items>
       <tns:DeclaredValue>$DeclaredValueToUse</tns:DeclaredValue>
       </tns:request>
   </tns:GetRateQuote>
 </soap:Body>
</soap:Envelope>
"@

       $headers = @{
           "Content-Type" = "text/xml; charset=utf-8"
           "SOAPAction" = "`"$soapAction`"" 
       }

       Write-Verbose "Calling RL API: Tariff $tariffNameForLog"
       try {
           $ProgressPreference = 'SilentlyContinue' 
           $response = Invoke-WebRequest -Uri $soapEndpoint -Method Post -Headers $headers -Body $soapRequestBody -UseBasicParsing -ErrorAction Stop
           [xml]$responseXml = $response.Content
           $nsManager = New-Object System.Xml.XmlNamespaceManager($responseXml.NameTable)
           $nsManager.AddNamespace("soap", $soapNamespace)
           $nsManager.AddNamespace("rl", $tnsNamespace) 
           $faultNode = $responseXml.SelectSingleNode("/soap:Envelope/soap:Body/soap:Fault", $nsManager)
           if ($faultNode) {
                $faultString = $faultNode.SelectSingleNode("faultstring", $nsManager).InnerText
                $faultCode = $faultNode.SelectSingleNode("faultcode", $nsManager).InnerText
                Write-Warning "RL API returned SOAP Fault for Tariff $($tariffNameForLog): Code='$faultCode', String='$faultString'" 
                Write-Verbose "RL Response XML (Fault - Tariff $tariffNameForLog): $($response.Content)"
                return $null
           }
           $rateQuoteResult = $responseXml.SelectSingleNode("/soap:Envelope/soap:Body/rl:GetRateQuoteResponse/rl:GetRateQuoteResult", $nsManager)
           if ($rateQuoteResult) {
               $quoteDetails = $rateQuoteResult.SelectSingleNode("rl:Result", $nsManager)
               if ($quoteDetails) {
                   $netChargeEntry = $quoteDetails.SelectSingleNode("rl:Charges/rl:Charge[rl:Type='NET']/rl:Amount", $nsManager)
                   if ($netChargeEntry) {
                       $totalChargeValue = $netChargeEntry.InnerText
                       try {
                           $cleanedRate = $totalChargeValue -replace '[$,]' 
                           $decimalRate = [decimal]$cleanedRate
                           Write-Verbose "RL OK: Tariff $tariffNameForLog Rate: $decimalRate"
                           return $decimalRate
                       } catch {
                           Write-Warning "RL Convert Fail for Tariff $($tariffNameForLog): Cannot convert rate '$totalChargeValue' to decimal. Error: $($_.Exception.Message)" 
                           return $null
                       }
                   } else {
                       $errorMsgNode = $quoteDetails.SelectSingleNode("rl:Errors/rl:string", $nsManager)
                       if($errorMsgNode){
                            Write-Warning "RL API Error in Response for Tariff $($tariffNameForLog): $($errorMsgNode.InnerText)" 
                       } else {
                            Write-Warning "RL Resp Missing 'NET' charge for Tariff $tariffNameForLog. Check response XML."
                       }
                       Write-Verbose "RL Response XML (No NET Charge - Tariff $tariffNameForLog): $($response.Content)"
                       return $null
                   }
               } else {
                   Write-Warning "RL Resp structure unexpected. Cannot find 'Result' element for Tariff $tariffNameForLog."
                   Write-Verbose "RL Response XML (Tariff $tariffNameForLog): $($response.Content)"
                   return $null
               }
           } else {
                Write-Warning "RL Resp structure unexpected. Cannot find 'GetRateQuoteResult' or 'Fault' element for Tariff $tariffNameForLog."
                Write-Verbose "RL Response XML (Tariff $tariffNameForLog): $($response.Content)"
                return $null
           }
       } catch {
           $errMsg = $_.Exception.Message; $statusCode = "N/A"; $eBody = "N/A"
           if ($_.Exception.Response) {
               try {$statusCode = $_.Exception.Response.StatusCode.value__} catch{}
               try {
                   $stream = $_.Exception.Response.GetResponseStream(); $reader = New-Object System.IO.StreamReader($stream); $eBody = $reader.ReadToEnd(); $reader.Close(); $stream.Close()
               } catch {$eBody="(Err reading resp body: $($_.Exception.Message))"}
           }
           $truncatedBody = if ($eBody.Length -gt 500) { $eBody.Substring(0, 500) + "..." } else { $eBody }
           $fullErrMsg = "RL FAIL: Tariff $tariffNameForLog. Error: $errMsg (HTTP $statusCode) Resp: $truncatedBody"
           Write-Warning $fullErrMsg; return $null
       }
   } 

Write-Verbose "TMS R+L Helper Functions loaded."
