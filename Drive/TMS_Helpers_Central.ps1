# TMS_Helpers_Central.ps1
# Description: Contains helper functions specific to Central Transport carrier operations,
#              including data normalization and API interaction.
#              This file should be dot-sourced by the main script(s) after TMS_Config.ps1.

# Assumes config variables like $script:centralApiUri are available from TMS_Config.ps1
# Assumes general helper functions (if any were used by these) are available.

# --- Data Normalization Functions ---

function Load-And-Normalize-CentralData {
    param([Parameter(Mandatory)][string]$CsvPath)
    Write-Host "`nLoading Central data: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan 
    $reqCols = @("Origin Postal Code", "Destination Postal Code", "Total Weight", "Freight Class 1")
    try {
        if (-not (Test-Path -Path $CsvPath -PathType Leaf)) {
            Write-Error "CSV file not found at '$CsvPath'."
            return $null
        }
        $rawData = Import-Csv -Path $CsvPath -ErrorAction Stop
        Write-Host " -> Rows read from CSV: $($rawData.Count)." -ForegroundColor Gray 
        if ($rawData.Count -eq 0) { Write-Warning "CSV empty."; return @() } # Return empty array for no data
        
        $headers = $rawData[0].PSObject.Properties.Name
        $missing = $reqCols | Where-Object { $_ -notin $headers } 
        if ($missing.Count -gt 0) { Write-Error "CSV missing Central columns: $($missing -join ', ')"; return $null }
        
        $normData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing Central data..." -ForegroundColor Gray
        $invalid = 0; $rowNum = 1 
        foreach ($row in $rawData) {
            $rowNum++
            $oZipRaw=$row."Origin Postal Code"; $dZipRaw=$row."Destination Postal Code"; $wtStrRaw=$row."Total Weight"; $clStrRaw=$row."Freight Class 1"
            $oZip=$oZipRaw.Trim(); $dZip=$dZipRaw.Trim(); $wtStr=$wtStrRaw.Trim(); $clStr=$clStrRaw.Trim(); $wtNum=$null
            $skipRow = $false
            
            if ([string]::IsNullOrWhiteSpace($oZip) -or $oZip.Length -lt 5) { $invalid++; Write-Verbose "Skip CTX Row ${rowNum}: Bad Origin Zip '$oZipRaw'"; $skipRow = $true }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($dZip) -or $dZip.Length -lt 5)) { $invalid++; Write-Verbose "Skip CTX Row ${rowNum}: Bad Dest Zip '$dZipRaw'"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($clStr)) { $invalid++; Write-Verbose "Skip CTX Row ${rowNum}: Bad Class '$clStrRaw'"; $skipRow = $true }
            if (-not $skipRow) { 
                try { 
                    $wtNum = [decimal]$wtStr
                    if($wtNum -le 0){throw "Weight must be positive."} 
                } catch { 
                    $invalid++; Write-Verbose "Skip CTX Row ${rowNum}: Bad Weight '$wtStrRaw' Error: $($_.Exception.Message)"; $skipRow = $true 
                } 
            }

            if (-not $skipRow) {
                 $normData.Add([PSCustomObject]@{
                    "Origin Postal Code" = $oZip 
                    "Destination Postal Code" = $dZip
                    "Total Weight" = $wtNum 
                    "Freight Class 1" = $clStr
                 })
            }
        }
        if ($invalid -gt 0) { Write-Warning " -> Skipped $invalid Central rows (missing/invalid data)." }
        Write-Host " -> OK: $($normData.Count) Central rows normalized." -ForegroundColor Green
        return $normData 
    } catch { 
        Write-Error "Error processing Central CSV '$CsvPath': $($_.Exception.Message)"
        return $null 
    }
}

# --- API Call Functions ---

function Invoke-CentralTransportApi {
    [CmdletBinding(DefaultParameterSetName = 'FromShipmentObject')]
    param(
        # Parameter Set 'FromShipmentObject'
        [Parameter(Mandatory, ParameterSetName='FromShipmentObject')]
        [PSCustomObject]$ShipmentData,

        [Parameter(Mandatory, ParameterSetName='FromShipmentObject')]
        [hashtable]$KeyData, 

        # Parameter Set 'FromIndividualParams'
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')]
        [string]$ApiKey, 
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')]
        [string]$OriginZip,
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')]
        [string]$DestinationZip,
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')]
        [double]$Weight,
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')]
        [string]$FreightClass,
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')]
        [string]$customerNumber,
        [Parameter(ParameterSetName='FromIndividualParams')]
        [string]$Accessorials = $null
    )
    
    $accessCodeToUse = $null; $customerNumberToUse = $null; $originZipToUse = $null; $destZipToUse = $null; $weightToUse = $null; $classToUse = $null; $tariffNameForLog = "Unknown"
    $localKeyData = $null 

    if ($PSCmdlet.ParameterSetName -eq 'FromShipmentObject') {
        $localKeyData = $KeyData 
        try {
            $originZipToUse = $ShipmentData.'Origin Postal Code'
            $destZipToUse = $ShipmentData.'Destination Postal Code'
            $weightToUse = [decimal]$ShipmentData.'Total Weight' 
            $classToUse = [string]$ShipmentData.'Freight Class 1'

            if ($localKeyData.ContainsKey('accessCode')) { $accessCodeToUse = $localKeyData.accessCode } else { throw "'accessCode' missing from KeyData." }
            if ($localKeyData.ContainsKey('customerNumber')) { $customerNumberToUse = $localKeyData.customerNumber } else { throw "'customerNumber' missing from KeyData." }
            $tariffNameForLog = if ($localKeyData.ContainsKey('TariffFileName')) { $localKeyData.TariffFileName } else { "UnknownTariff" }
            if (-not $localKeyData.ContainsKey('Name')){$localKeyData.Name=$tariffNameForLog} 
            if ([string]::IsNullOrWhiteSpace($accessCodeToUse) -or [string]::IsNullOrWhiteSpace($customerNumberToUse)) { throw "Credentials empty in KeyData."}
        } catch {
            Write-Warning "Central Extract Fail from Shipment Object (Tariff: $($tariffNameForLog)): $($_.Exception.Message)"
            return $null
        }
    } elseif ($PSCmdlet.ParameterSetName -eq 'FromIndividualParams') {
         $accessCodeToUse = $ApiKey
         $customerNumberToUse = $customerNumber 
         $originZipToUse = $OriginZip
         $destZipToUse = $DestinationZip
         $weightToUse = [decimal]$Weight 
         $classToUse = $FreightClass
         $tariffNameForLog = "SingleQuoteCall"
         if ($PSBoundParameters.ContainsKey('KeyData')) {
             $localKeyData = $PSBoundParameters['KeyData']
         } else {
             $localKeyData = @{Name=$tariffNameForLog; TariffFileName=$tariffNameForLog}
         }
         if(-not $localKeyData.ContainsKey('TariffFileName')){$localKeyData.TariffFileName = $tariffNameForLog}
    } else {
        Write-Error "CTX Internal Error: Invalid parameter set used ('$($PSCmdlet.ParameterSetName)')."
        return $null
    }

    $ctxMissing = @()
    if ([string]::IsNullOrWhiteSpace($originZipToUse)) { $ctxMissing += "OriginZip" }
    if ([string]::IsNullOrWhiteSpace($destZipToUse)) { $ctxMissing += "DestinationZip" }
    if ($null -eq $weightToUse -or $weightToUse -le 0) { $ctxMissing += "Weight(<=0 or invalid: '$($weightToUse)')" } 
    if ([string]::IsNullOrWhiteSpace($classToUse)) { $ctxMissing += "Class" }
    if ([string]::IsNullOrWhiteSpace($accessCodeToUse)) { $ctxMissing += "AccessCode(ApiKey)" }
    if ([string]::IsNullOrWhiteSpace($customerNumberToUse)) { $ctxMissing += "CustomerNumber" } 

    if ($ctxMissing.Count -gt 0) {
        $logTariffName = if($localKeyData -and $localKeyData.ContainsKey('TariffFileName')){$localKeyData.TariffFileName}else{$tariffNameForLog}
        Write-Warning "CTX Skip: Tariff '$($logTariffName)' - Missing required data: $($ctxMissing -join ', ')."
        return $null
    }

    $rateItemsArray = @( @{ id = 1; weight = $weightToUse; itemClass = $classToUse } )
    $payload = @{
        accessCode = $accessCodeToUse;
        request = @{
            originZipCode = $originZipToUse;
            destinationZipCode = $destZipToUse;
            customerNumber = [string]$customerNumberToUse; 
            pickupDate = (Get-Date -Format 'MM/dd/yyyy');
            customerRole = "shipper";
            rateItems = $rateItemsArray;
            useDefaultTariff = $false
        }
    } | ConvertTo-Json -Depth 5

    $headers = @{ 'Content-Type' = 'application/json' }
    $logTariffNameCall = if($localKeyData -and $localKeyData.ContainsKey('TariffFileName')){$localKeyData.TariffFileName}else{$tariffNameForLog}
    Write-Verbose "Calling CTX API: Tariff $($logTariffNameCall)"

    try {
        $apiUrl = $script:centralApiUri 
        if ([string]::IsNullOrWhiteSpace($apiUrl)) { throw "Central API URI ('$($apiUrl)') is not defined or empty."}
        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $payload -ErrorAction Stop
        Write-Verbose "CTX OK: Tariff $($logTariffNameCall)"
        $totalChargeValue = $null
        if ($response -ne $null) {
            if ($response.PSObject.Properties.Name -contains 'rateTotal') {
                $totalChargeValue = $response.rateTotal
            }
        }
        if ($totalChargeValue -ne $null) {
            try {
                 $cleanedRate = $totalChargeValue -replace '[$,]' 
                 $decimalRate = [decimal]$cleanedRate
                 return $decimalRate
            } catch {
                 Write-Warning "CTX Convert Fail for Tariff $($logTariffNameCall): Cannot convert rate '$totalChargeValue' to decimal. Error: $($_.Exception.Message)"
                 return $null
            }
        } else {
             Write-Warning "CTX Resp missing 'rateTotal': Tariff $($logTariffNameCall)"
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
         $fullErrMsg = "CTX FAIL: Tariff $($logTariffNameCall). Error: $errMsg (HTTP $statusCode) Resp: $truncatedBody"
         Write-Warning $fullErrMsg;
         return $null
    }
}

Write-Verbose "TMS Central Transport Helper Functions loaded."
