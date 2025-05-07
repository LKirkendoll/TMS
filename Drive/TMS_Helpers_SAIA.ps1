# TMS_Helpers_SAIA.ps1
# Description: Contains helper functions specific to SAIA carrier operations,
#              including data normalization and API interaction.
#              This file should be dot-sourced by the main script(s) after TMS_Config.ps1.

# Assumes config variables like $script:saiaApiUri are available from TMS_Config.ps1
# Assumes general helper functions (if any were used by these) are available.

# --- Data Normalization Functions ---

function Load-And-Normalize-SAIAData {
    param([Parameter(Mandatory)][string]$CsvPath)
    Write-Host "`nLoading SAIA data: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan
    $reqCols = @( "Origin Postal Code", "Destination Postal Code", "Total Weight", "Freight Class 1", "Origin City", "Origin State", "Destination City", "Destination State") # Base required
    $optCols = @( "Total Units", "Total Density" ) # Optional for dims
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
        if ($missing.Count -gt 0) { Write-Error "CSV missing required SAIA columns: $($missing -join ', ')"; return $null }
        
        $missingOpt = $optCols | Where-Object { $_ -notin $headers }
        if($missingOpt.Count -gt 0){ Write-Warning "CSV missing optional SAIA columns used for dimension calculation: $($missingOpt -join ', ')" }

        $normData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing SAIA data..." -ForegroundColor Gray
        $invalid = 0; $rowNum = 1
        foreach ($row in $rawData) {
            $rowNum++
            # Read values, trimming whitespace
            $oZipRaw=$row."Origin Postal Code"; $dZipRaw=$row."Destination Postal Code"; $wtStrRaw=$row."Total Weight"; $clStrRaw=$row."Freight Class 1"; $oCityRaw=$row."Origin City"; $oStateRaw=$row."Origin State"; $dCityRaw=$row."Destination City"; $dStateRaw=$row."Destination State"

            $oZip=$oZipRaw.Trim(); $dZip=$dZipRaw.Trim(); $wtStr=$wtStrRaw.Trim(); $clStr=$clStrRaw.Trim(); $oCity=$oCityRaw.Trim(); $oState=$oStateRaw.Trim(); $dCity=$dCityRaw.Trim(); $dState=$dStateRaw.Trim(); $wtNum=$null; $unitsNum=$null; $densityNum=$null

            $skipRow = $false
            if ([string]::IsNullOrWhiteSpace($oZip) -or $oZip.Length -lt 5) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Origin Zip '$oZipRaw'"; $skipRow = $true }
            if (-not $skipRow -and ([string]::IsNullOrWhiteSpace($dZip) -or $dZip.Length -lt 5)) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Dest Zip '$dZipRaw'"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($clStr)) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Class '$clStrRaw'"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($oCity)) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Origin City '$oCityRaw'"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($oState)) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Origin State '$oStateRaw'"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($dCity)) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Dest City '$dCityRaw'"; $skipRow = $true }
            if (-not $skipRow -and [string]::IsNullOrWhiteSpace($dState)) { $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Dest State '$dStateRaw'"; $skipRow = $true }
            if (-not $skipRow) { 
                try { 
                    $wtNum = [decimal]$wtStr
                    if($wtNum -le 0){throw "Weight must be positive."} 
                } catch { 
                    $invalid++; Write-Verbose "Skip SAIA Row ${rowNum}: Bad Weight '$wtStrRaw' Error: $($_.Exception.Message)"; $skipRow = $true 
                } 
            }

            if ($skipRow) { continue }

            if ($headers -contains 'Total Units') { try { $unitsNum = [int]$row.'Total Units'.Trim(); if ($unitsNum -le 0) {$unitsNum = $null} } catch { Write-Verbose "Invalid Total Units in row ${rowNum}: $($row.'Total Units')"} }
            if ($headers -contains 'Total Density') { try { $densityNum = [double]$row.'Total Density'.Trim(); if ($densityNum -le 0) {$densityNum = $null} } catch { Write-Verbose "Invalid Total Density in row ${rowNum}: $($row.'Total Density')"} }

            $detailProps = [ordered]@{
                weight = $wtNum 
                class = $clStr 
                length = 1.0 
                width = 1.0
                height = 1.0
                units = 1 
            }
            if ($unitsNum -ne $null -and $unitsNum -gt 0) { $detailProps.units = $unitsNum }
            $detailObject = [PSCustomObject]$detailProps

            $newRow = [PSCustomObject]@{
                OriginZip        = $oZip
                DestinationZip   = $dZip
                OriginCity       = $oCity
                OriginState      = $oState
                DestinationCity  = $dCity
                DestinationState = $dState
                details          = @($detailObject) 
                'Total Weight'   = $wtNum 
                'Total Units'    = $unitsNum
                'Total Density'  = $densityNum
                'Freight Class 1' = $clStr 
            }
            $normData.Add($newRow) 
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
        [Parameter(Mandatory=$true)] [decimal]$Weight, 
        [Parameter(Mandatory=$true)] [string]$Class,   
        [Parameter(Mandatory=$false)] [object]$Details = $null 
    )

    $saiaUserID = $null; $saiaPassword = $null; $saiaRQKey = $null;
    $accountCodeToUse = $null 
    $detailsArray = @()
    $tariffNameForLog = if ($KeyData.ContainsKey('TariffFileName')) { $KeyData.TariffFileName } else { "UnknownTariff" }
    if (-not $KeyData.ContainsKey('Name')) { $KeyData.Name = $tariffNameForLog } 

    try {
        if ($KeyData.ContainsKey('UserID')) { $saiaUserID = $KeyData.UserID } else { Write-Verbose "UserID missing from KeyData for Acct '$tariffNameForLog'."}
        if ($KeyData.ContainsKey('Password')) { $saiaPassword = $KeyData.Password } else { Write-Verbose "Password missing from KeyData for Acct '$tariffNameForLog'."}
        if ($KeyData.ContainsKey('RQKey')) { $saiaRQKey = $KeyData.RQKey } else { Write-Verbose "RQKey missing from KeyData for Acct '$tariffNameForLog'."}
        if ($KeyData.ContainsKey('AccountCode')) { $accountCodeToUse = $KeyData.AccountCode } else { Write-Verbose "AccountCode not found in KeyData for SAIA Acct '$tariffNameForLog'." }

        if ($Weight -gt 0 -and -not [string]::IsNullOrWhiteSpace($Class)) {
            $itemWeight = 0; $itemClass = 0.0; $itemUnits = 1; $itemLength = 1.0; $itemWidth = 1.0; $itemHeight = 1.0; 
            try { $itemWeight = [int]$Weight } catch { Write-Warning "Cannot convert weight '$Weight' to int for details for Acct '$tariffNameForLog'." }
            if (-not [double]::TryParse($Class, [ref]$itemClass)) { Write-Warning "Could not parse Class '$Class' as a double for details for Acct '$tariffNameForLog'. Using 0.0." }

            if ($null -ne $Details) {
                 $detailItemToParse = $null
                 if ($Details -is [array] -and $Details.Count -gt 0) {
                     $detailItemToParse = $Details[0]
                 } elseif ($Details -is [PSCustomObject]) {
                     $detailItemToParse = $Details
                 }
                 if ($null -ne $detailItemToParse) {
                     if($detailItemToParse.PSObject.Properties.Name -contains 'units' -and $null -ne $detailItemToParse.units) { try { $itemUnits = [int]$detailItemToParse.units; if($itemUnits -le 0) {$itemUnits = 1} } catch { $itemUnits=1} }
                     if($detailItemToParse.PSObject.Properties.Name -contains 'length' -and $null -ne $detailItemToParse.length) { try { $itemLength = [double]$detailItemToParse.length } catch {$itemLength=1.0} }
                     if($detailItemToParse.PSObject.Properties.Name -contains 'width' -and $null -ne $detailItemToParse.width) { try { $itemWidth = [double]$detailItemToParse.width } catch {$itemWidth=1.0} }
                     if($detailItemToParse.PSObject.Properties.Name -contains 'height' -and $null -ne $detailItemToParse.height) { try { $itemHeight = [double]$detailItemToParse.height } catch {$itemHeight=1.0} }
                 } else {
                     Write-Verbose "Passed Details object was null or empty array for Acct '$tariffNameForLog'. Using default dims/units."
                 }
            }
            if ($itemWeight -gt 0) { 
                $detailsArray += @{ length=$itemLength; width=$itemWidth; height=$itemHeight; weight=$itemWeight; class=$itemClass; units=$itemUnits }
            } else {
                Write-Warning "Invalid integer weight ($itemWeight) derived from '$Weight' for Acct '$tariffNameForLog'. Details array might be empty or incomplete."
            }
        } else {
            Write-Warning "Invalid weight ($Weight) or class ('$Class' is NullOrWhitespace: $([string]::IsNullOrWhiteSpace($Class))) for Acct '$tariffNameForLog'. Details array will be empty."
        }
    } catch { Write-Warning "SAIA Extract/Prep Fail for Acct '$tariffNameForLog': $($_.Exception.Message)"; return $null } 


    $missingFields = @()
    if ([string]::IsNullOrWhiteSpace($OriginZip)) { $missingFields += "OriginZip" }
    if ([string]::IsNullOrWhiteSpace($DestinationZip)) { $missingFields += "DestinationZip" }
    if ($null -eq $Weight -or $Weight -le 0) { $missingFields += "Weight(<=0 or invalid: '$Weight')" }
    if ([string]::IsNullOrWhiteSpace($Class)) { $missingFields += "Class" }
    if ($detailsArray.Count -eq 0) { $missingFields += "Details(Commodity Data/Invalid Class/Wt)" }
    if ([string]::IsNullOrWhiteSpace($saiaRQKey) -and ([string]::IsNullOrWhiteSpace($saiaUserID) -or [string]::IsNullOrWhiteSpace($saiaPassword))) {
        $missingFields += "Credentials (Missing RQKey, AND missing UserID/Password pair)"
    }
    if ([string]::IsNullOrWhiteSpace($OriginCity)) { $missingFields += "OriginCity" }
    if ([string]::IsNullOrWhiteSpace($OriginState)) { $missingFields += "OriginState" }
    if ([string]::IsNullOrWhiteSpace($DestinationCity)) { $missingFields += "DestinationCity" }
    if ([string]::IsNullOrWhiteSpace($DestinationState)) { $missingFields += "DestinationState" }

    if ($missingFields.Count -gt 0) {
        $contextZips = ""
        if (-not [string]::IsNullOrWhiteSpace($OriginZip)) { $contextZips += " OZip:$OriginZip" }
        if (-not [string]::IsNullOrWhiteSpace($DestinationZip)) { $contextZips += " DZip:$DestinationZip" }
        Write-Warning "SAIA Skip: Acct '$tariffNameForLog'$contextZips - Missing required data: $($missingFields -join ', ')."
        return $null
    }

    $calculatedTotalCube = 0.0
    try {
        $totalVolumeInches = 0.0
        foreach($item in $detailsArray) {
            if ($item.PSObject.Properties.Name -contains 'length' -and $item.length -is [double] -and $item.length -gt 0 -and
                $item.PSObject.Properties.Name -contains 'width' -and $item.width -is [double] -and $item.width -gt 0 -and
                $item.PSObject.Properties.Name -contains 'height' -and $item.height -is [double] -and $item.height -gt 0 -and
                $item.PSObject.Properties.Name -contains 'units' -and $item.units -is [int] -and $item.units -gt 0) {
                 $totalVolumeInches += ($item.length * $item.width * $item.height * $item.units)
            } else {
                 Write-Warning "Skipping item in cube calculation due to missing, zero, or non-numeric dimension/unit for Acct '$tariffNameForLog'."
            }
        }
        if ($totalVolumeInches -gt 0) {
             $calculatedTotalCube = [Math]::Round($totalVolumeInches / 1728, 2) 
        }
    } catch {
        Write-Warning "Could not calculate totalCube for Acct '$tariffNameForLog': $($_.Exception.Message)"
        $calculatedTotalCube = 0.0
    }

    $payloadObject = [ordered]@{ 
        userID = $saiaUserID
        password = $saiaPassword
        payer = "Shipper"
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
        details = $detailsArray
    }

    if (-not [string]::IsNullOrWhiteSpace($accountCodeToUse)) {
        if ($null -eq $payloadObject.origin) { $payloadObject.origin = @{} }
        if ($null -eq $payloadObject.destination) { $payloadObject.destination = @{} }
        $payloadObject.origin.accountCode = $accountCodeToUse
        $payloadObject.destination.accountCode = $accountCodeToUse
    }

    if ($payloadObject.Keys -contains 'userID' -and [string]::IsNullOrWhiteSpace($payloadObject['userID'])) { $payloadObject.Remove('userID') }
    if ($payloadObject.Keys -contains 'password' -and [string]::IsNullOrWhiteSpace($payloadObject['password'])) { $payloadObject.Remove('password') }

    $payload = $payloadObject | ConvertTo-Json -Depth 10

    $headers = @{ 'Content-Type' = 'application/json'; 'Cache-Control' = 'no-cache' }
    if (-not [string]::IsNullOrWhiteSpace($saiaRQKey)) {
        $headers.'RQ-Key' = $saiaRQKey
    }

    try {
        $apiUrl = $script:saiaApiUri
        if ([string]::IsNullOrWhiteSpace($apiUrl)) { throw "SAIA API URI ($($apiUrl)) is not defined or empty."}

        if (-not $headers.ContainsKey('RQ-Key') -and ([string]::IsNullOrWhiteSpace($saiaUserID) -or [string]::IsNullOrWhiteSpace($saiaPassword)) ) {
             throw "Cannot call SAIA API for Acct '$tariffNameForLog': No RQ-Key header provided AND UserID/Password pair is incomplete."
        }

        $response = Invoke-RestMethod -Uri $apiUrl -Method Post -Headers $headers -Body $payload -ErrorAction Stop
        Write-Verbose "SAIA OK: Acct '$tariffNameForLog'" 

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
                 return $decimalRate
            } catch {
                 Write-Warning "SAIA Convert Fail for Acct '$tariffNameForLog': Cannot convert rate '$totalChargeValue' to decimal. Error: $($_.Exception.Message)"; return $null
            }
        } else {
             Write-Warning "SAIA Resp missing 'rateDetails.totalInvoice' or structure invalid for Acct '$tariffNameForLog'.";
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
        $fullErrMsg = "SAIA FAIL: Acct '$tariffNameForLog'. Error: $errMsg (HTTP $statusCode) Resp: $truncatedBody"
        Write-Warning $fullErrMsg; 
        return $null
     }
} 

Write-Verbose "TMS SAIA Helper Functions loaded."
