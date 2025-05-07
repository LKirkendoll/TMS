# TMS_Helpers_General.ps1
# Description: Contains general reusable helper functions for the TMS Tool.
#              This file should be dot-sourced by the main script(s).

# Ensure necessary assemblies are loaded (if any GUI elements are directly created here, though most are in TMS_GUI.ps1)
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
} catch {
    Write-Error "Failed to load required .NET Assembly: System.Windows.Forms. Ensure .NET Framework is available."
    throw "Assembly load failed."
}

# --- File/Folder Functions ---

function Ensure-DirectoryExists {
    param( [Parameter(Mandatory=$true)][string]$Path )
    if (-not (Test-Path -Path $Path -PathType Container)) {
        Write-Warning "Required folder '$(Split-Path -Path $Path -Leaf)' ('$Path') not found. Creating it..."
        try {
            New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
        } catch {
            Write-Error "Failed to create folder '$Path': $($_.Exception.Message)"; throw "Directory creation failed."
        }
    }
}

# --- User Interface Functions (Primarily for Console, but Select-CsvFile is for GUI) ---

function Select-CsvFile {
    param( [string]$DialogTitle = "Select CSV File", [string]$InitialDirectory = $script:shipmentDataFolderPath )
    if ([string]::IsNullOrWhiteSpace($InitialDirectory) -or -not (Test-Path $InitialDirectory)) {
         $InitialDirectory = $script:scriptRoot # Assumes $script:scriptRoot is set by the calling GUI script
         if ($null -eq $InitialDirectory -or -not(Test-Path $InitialDirectory)) {
              Write-Warning "Cannot determine initial directory from script:scriptRoot. Defaulting to C:\."
              $InitialDirectory = "C:\"
         }
    }

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $DialogTitle
    $dialog.InitialDirectory = $InitialDirectory
    $dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.FileName }
    else { Write-Warning "File selection cancelled."; return $null }
}

function Clear-HostAndDrawHeader { # Console Specific
    param(
        [Parameter(Mandatory=$true)][string]$Title,
        [string]$User = $null 
    )
    Clear-Host
    $border = "=" * ($Title.Length + 4) 
    Write-Host $border -ForegroundColor Blue
    Write-Host ("  " + $Title + "  ") -ForegroundColor White
    if (-not [string]::IsNullOrWhiteSpace($User)) {
        $userDisplay = "  User: " + $User
        if ($userDisplay.Length -gt ($border.Length - 2)) {
            $userDisplay = $userDisplay.Substring(0, $border.Length - 5) + "..."
        }
         Write-Host $userDisplay.PadRight($border.Length - 2) -ForegroundColor White
    }
    Write-Host $border -ForegroundColor Blue
    Write-Host "" 
}

function Write-LoadingBar { # Console Specific
    param( [Parameter(Mandatory=$true)][int]$PercentComplete, [string]$Message = "Processing..." )
    $validPercent = $PercentComplete
    if ($validPercent -lt 0) { $validPercent = 0 }
    if ($validPercent -gt 100) { $validPercent = 100 }
    Write-Progress -Activity $Message -Status "$validPercent% Complete" -PercentComplete $validPercent
}

function Show-Highlights { # Console Specific
    param()
    while ($true) {
        $choice = Read-Host "Show highlights in console during processing? (Y/N)"
        if ($choice -match '^[Yy]$') { return $true }
        if ($choice -match '^[Nn]$') { return $false }
        Write-Warning "Invalid input. Please enter Y or N."
    }
}

function Open-FileExplorer {
    param( [Parameter(Mandatory=$true)][string]$Path )
    if (Test-Path -Path $Path) {
        try {
             Invoke-Item -Path $Path -ErrorAction Stop
        } catch {
             Write-Error "Failed to open path '$Path' in File Explorer: $($_.Exception.Message)"
        }
    }
    else { Write-Warning "Cannot open path in File Explorer because it does not exist: $Path" }
}

# --- Data Handling Functions (General) ---

function Load-KeysFromFolder { # Used by GUI startup
    param( [Parameter(Mandatory)][string]$KeysFolderPath, [Parameter(Mandatory)][string]$CarrierName )
    $loadedKeysAndMargins = @{}; Write-Verbose "Loading keys/margins for $CarrierName from: $KeysFolderPath (Manual Parse)"
    if (-not (Test-Path -Path $KeysFolderPath -PathType Container)) { Write-Warning "Key folder '$KeysFolderPath' not found."; return $loadedKeysAndMargins }
    $keyFiles = Get-ChildItem -Path $KeysFolderPath -Filter "*.txt" -File -ErrorAction SilentlyContinue
    if ($keyFiles) {
        foreach ($file in $keyFiles) {
            $keyNameFromFile = $file.BaseName; $keyDataHashtable = @{}; $isFirstLine = $true
            try {
                $lines = Get-Content -Path $file.FullName -ErrorAction Stop
                foreach ($line in $lines) {
                    $trimmedLine = $line.Trim(); if ([string]::IsNullOrWhiteSpace($trimmedLine) -or $trimmedLine.StartsWith('#')) { continue } 
                    $equalsIndex = $trimmedLine.IndexOf('=')
                    if ($equalsIndex -gt 0) {
                        $key = $trimmedLine.Substring(0, $equalsIndex).Trim(); $value = $trimmedLine.Substring($equalsIndex + 1).Trim()
                        # Strip potential source control markers from the key name on the first line (Corrected regex)
                        if ($isFirstLine -and $key -match '^\[source:\s*\d+\]\s*(.*)') { $actualKey = $Matches[1].Trim(); if (-not [string]::IsNullOrEmpty($actualKey)) { $key = $actualKey } }
                        $isFirstLine = $false
                        if (-not [string]::IsNullOrEmpty($key)) { $keyDataHashtable[$key] = $value }
                    } else { Write-Warning "Skipping line (no '=') in '$($file.Name)': $line" }
                } 
                if ($keyDataHashtable.Count -gt 0) {
                    $keyDataHashtable['TariffFileName'] = $keyNameFromFile 
                    if (-not $keyDataHashtable.ContainsKey('Name')) { $keyDataHashtable['Name'] = $keyNameFromFile }
                    $loadedKeysAndMargins[$keyNameFromFile] = $keyDataHashtable; Write-Verbose "Loaded data for '$keyNameFromFile'."

                    # Carrier Specific Validation (Moved to carrier-specific helpers if needed, or kept generic here)
                    if (-not $keyDataHashtable.ContainsKey('MarginPercent')) { Write-Warning "'$($file.Name)' missing 'MarginPercent'."} elseif ($keyDataHashtable['MarginPercent'] -as [double] -eq $null) { Write-Warning "Invalid 'MarginPercent' in '$($file.Name)'." }
                    # Basic checks that could apply to any key file
                    # Add more general checks if applicable
                } else { Write-Warning "No valid Key=Value pairs in '$($file.Name)'." }
            } catch { Write-Warning "Could not process key file '$($file.Name)': $($_.Exception.Message)" }
        } 
    } else { Write-Verbose "No .txt key files found in '$KeysFolderPath'." }
    Write-Host "Loaded $($loadedKeysAndMargins.Count) $CarrierName key(s)/account(s)." -ForegroundColor Gray
    return $loadedKeysAndMargins
}


function Get-PermittedKeys { # Used by GUI
    param( [Parameter(Mandatory)][hashtable]$AllKeys, [Parameter(Mandatory)][array]$AllowedKeyNames )
    $permittedKeys = @{}
    if ($null -ne $AllowedKeyNames) {
        foreach ($allowedName in $AllowedKeyNames) {
             if ([string]::IsNullOrWhiteSpace($allowedName)) { continue }
             if ($AllKeys.ContainsKey($allowedName)) {
                 if ($AllKeys[$allowedName] -is [hashtable]) { $permittedKeys[$allowedName] = $AllKeys[$allowedName] }
                 else { Write-Warning "Value for key '$allowedName' was not a hashtable." }
             } else { Write-Warning "User allowed key '$allowedName' not found in loaded keys for the current carrier." }
        }
    }
    return $permittedKeys
}

function Identify-BlanketKey { # Used by console version, might be useful for GUI logic too
    param( [Parameter(Mandatory)][hashtable]$PermittedKeys )
    foreach ($keyName in $PermittedKeys.Keys) { if ($keyName -match 'Blanket') { Write-Verbose "Identified Blanket Key: $keyName"; return $PermittedKeys[$keyName] } }
    foreach ($keyDataValue in $PermittedKeys.Values) {
        if ($keyDataValue -is [hashtable]) {
            if (($keyDataValue.ContainsKey('Name') -and $keyDataValue.Name -match 'Blanket') -or
                ($keyDataValue.ContainsKey('AccountName') -and $keyDataValue.AccountName -match 'Blanket') -or
                ($keyDataValue.ContainsKey('TariffFileName') -and $keyDataValue.TariffFileName -match 'Blanket')) { 
                Write-Verbose "Identified Blanket Key via property/filename: $($keyDataValue.TariffFileName)"
                return $keyDataValue
            }
        }
    }
    Write-Warning "Could not identify 'Blanket' key among permitted keys."; return $null
}

function Select-SingleKeyEntry { # Console Specific
    param( [Parameter(Mandatory)][hashtable]$AvailableKeys, [Parameter(Mandatory)][string]$PromptMessage, [array]$ExcludeNames=@() )
    $selectableKeys = $AvailableKeys.Keys | Where-Object { $_ -notin $ExcludeNames } | Sort-Object
    if ($selectableKeys.Count -eq 0) { Write-Warning "No keys available matching criteria."; return $null }

    Write-Verbose "DEBUG (Select-SingleKeyEntry): Selectable keys array:"
    $selectableKeys | ForEach-Object { Write-Verbose "  - '$_'" }

    Write-Host "`n$PromptMessage" -ForegroundColor Yellow; Write-Host "---" -ForegroundColor DarkGray
    for ($i=0; $i -lt $selectableKeys.Count; $i++) {
        Write-Verbose "DEBUG (Select-SingleKeyEntry): Displaying item index {$i}: '$($selectableKeys[$i])'"
        Write-Host (" [{0,2}] : {1}" -f ($i + 1), $selectableKeys[$i])
    }
    Write-Host " [ b ] : Go Back" -ForegroundColor White; Write-Host "---" -ForegroundColor DarkGray
    $selectedDetails = $null
    while($true) { 
        $input = Read-Host "Enter number (1-$($selectableKeys.Count)) or 'b'"
        if ($input -eq 'b') { Write-Host "Cancelled." -ForegroundColor Yellow; return $null } 
        try {
            if ($input -match '^\d+$') {
                $idx = [int]$input 
                if ($idx -ge 1 -and $idx -le $selectableKeys.Count) { 
                    $selectedKeyName = $selectableKeys[$idx-1] 
                    $selectedDetails = $AvailableKeys[$selectedKeyName]
                    if ($selectedDetails -ne $null) {
                        Write-Host " -> Selected: '$selectedKeyName'" -ForegroundColor Green; Start-Sleep -Milliseconds 300
                        break 
                    } else {
                        Write-Warning "Internal error: Could not retrieve details for selected key '$selectedKeyName'."
                        $selectedDetails = $null 
                    }
                } else { Write-Warning "Out of range." }
            } else { Write-Warning "Invalid input."} 
        } catch {
             Write-Warning "Input error: $($_.Exception.Message)"; $selectedDetails = $null 
        }
    } 
    if ($selectedDetails -is [hashtable] -and -not $selectedDetails.ContainsKey('Name')) {
        $nameToAdd = if ($selectedDetails.ContainsKey('AccountName')) { $selectedDetails.AccountName } `
                     elseif ($selectedDetails.ContainsKey('TariffFileName')) { $selectedDetails.TariffFileName } `
                     else { $selectedKeyName } 
        if ([string]::IsNullOrWhiteSpace($nameToAdd) -and $selectedDetails.ContainsKey('TariffFileName')) {
            $nameToAdd = $selectedDetails.TariffFileName
        }
        $selectedDetails['Name'] = $nameToAdd
    }
    return $selectedDetails
}

# --- Quoting Logic Helpers (General) ---

function Get-MinimumRate {
    param(
        [Parameter(Mandatory=$true)][hashtable]$RateResults 
    )
    $lowestCost = $null
    $bestTariff = $null
    foreach ($tariffName in $RateResults.Keys) {
         $cost = $RateResults[$tariffName] 
         if ($cost -ne $null -and $cost -is [decimal] -and $cost -gt 0) { 
              if ($lowestCost -eq $null -or $cost -lt $lowestCost) {
                   $lowestCost = $cost
                   $bestTariff = $tariffName
              }
         }
    }
    if ($lowestCost -ne $null) {
        return [PSCustomObject]@{ TariffName = $bestTariff; Cost = $lowestCost }
    } else {
        return $null
    }
}

function Get-HistoricalAveragePrice {
    param( [Parameter(Mandatory)] [string]$OriginZip, [Parameter(Mandatory)] [string]$DestinationZip, [Parameter(Mandatory)] [double]$Weight, [Parameter(Mandatory)] [string]$FreightClass )
    $histFileName = $Global:HistoricalDataSourceFileName 
    if ([string]::IsNullOrWhiteSpace($histFileName)) { Write-Warning "HistoricalDataSourceFileName not set in config."; return $null }
    $histPath = Join-Path $script:shipmentDataFolderPath $histFileName # Assumes $script:shipmentDataFolderPath is set
    Write-Verbose "Lookup hist avg: $OriginZip->$DestinationZip Wt:$Weight Cls:$FreightClass File: $(Split-Path $histPath -Leaf)"
    $cutoff = (Get-Date).AddMonths(-12); $avgPrice = $null
    if (-not (Test-Path $histPath -PathType Leaf)) { Write-Warning "Hist file missing: $histPath"; return $null }
    if ([string]::IsNullOrWhiteSpace($OriginZip) -or $OriginZip.Length -lt 3 -or [string]::IsNullOrWhiteSpace($DestinationZip) -or $DestinationZip.Length -lt 3) { Write-Warning "Hist Lookup Skip: Origin/Dest ZIP too short or missing."; return $null }

    $oZip3=$OriginZip.Substring(0,3); $dZip3=$DestinationZip.Substring(0,3)
    $tolerance = $Global:HistoricalWeightTolerance 
    if ($null -eq $tolerance) { Write-Warning "HistoricalWeightTolerance not set in config. Defaulting to 0.10"; $tolerance = 0.10 }
    $minWt=$Weight*(1.0 - $tolerance); $maxWt=$Weight*(1.0 + $tolerance)
    Write-Verbose " -> Hist Weight Range: $minWt - $maxWt"

    try {
        $hist = Import-Csv -Path $histPath -ErrorAction Stop
        $oZipCol='Origin Postal Code'; $dZipCol='Destination Postal Code'; $wtCol='Total Weight'; $clsCol='Freight Class 1'; $prcCol='Price'; $dtCol='Booked Date' 
        $hdr = $hist[0].PSObject.Properties.Name
        $reqHistCols = @($oZipCol, $dZipCol, $wtCol, $clsCol, $prcCol, $dtCol)
        $missingHistCols = $reqHistCols | Where-Object { $_ -notin $hdr }
        if ($missingHistCols) { Write-Warning "Hist file '$histPath' missing required columns: $($missingHistCols -join ', ')"; return $null }

        $similarShipmentsData = $hist | Where-Object {
            $rowWeight = $null; $rowPrice = $null; $rowDate = $null; $rowOZip = $null; $rowDZip = $null; $rowClass = $null
            $isValid = $true
            try { $rowOZip = $_.$oZipCol; if([string]::IsNullOrWhiteSpace($rowOZip) -or $rowOZip.Length -lt 3) {$isValid=$false} } catch {$isValid=$false}
            if($isValid){ try { $rowDZip = $_.$dZipCol; if([string]::IsNullOrWhiteSpace($rowDZip) -or $rowDZip.Length -lt 3) {$isValid=$false} } catch {$isValid=$false} }
            if($isValid){ try { $rowClass = $_.$clsCol; if([string]::IsNullOrWhiteSpace($rowClass)) {$isValid=$false} } catch {$isValid=$false} }
            if($isValid){ try { $rowWeight = [double]$_.$wtCol; if($rowWeight -le 0) {$isValid=$false} } catch {$isValid=$false} }
            if($isValid){ try { $rowPrice = [double]$_.$prcCol; if($rowPrice -le 0) {$isValid=$false} } catch {$isValid=$false} }
            if($isValid){ try { $rowDate = [datetime]$_.$dtCol } catch {$isValid=$false} }

            if ($isValid) {
                ($rowDate -ge $cutoff) -and
                ($rowOZip.Substring(0, 3) -eq $oZip3) -and
                ($rowDZip.Substring(0, 3) -eq $dZip3) -and
                ($rowClass -eq $FreightClass) -and
                ($rowWeight -ge $minWt) -and
                ($rowWeight -le $maxWt)
            } else {
                $false 
            }
        } | Select-Object -ExpandProperty $prcCol 

        if ($similarShipmentsData -and $similarShipmentsData.Count -gt 0) {
            $measureResult = $similarShipmentsData | Measure-Object -Average -ErrorAction Stop
            if ($measureResult) {
                 if ($measureResult.PSObject.Properties.Name -contains 'Average') {
                     $avgPrice = [math]::Round($measureResult.Average, 2)
                     Write-Verbose "Found $($similarShipmentsData.Count) hist matches (avg price: $avgPrice)"
                 } else {
                      Write-Warning "Measure-Object result did not contain 'Average' property."
                 }
            } else { Write-Warning "Measure-Object failed for historical data." }
        } else { Write-Verbose "No similar hist found matching all criteria." }
    } catch { Write-Error "Hist proc err: $($_.Exception.Message)" }
    return $avgPrice
}

function Calculate-QuotePrice {
    param( [Parameter(Mandatory)] [decimal]$LowestCarrierCost, [Parameter(Mandatory)] [string]$OriginZip, [Parameter(Mandatory)] [string]$DestinationZip, [Parameter(Mandatory)] [double]$Weight, [Parameter(Mandatory)] [string]$FreightClass, [Parameter(Mandatory)] [double]$MarginPercent )
    $stdMarginPrice = $null
    $marginDec = [decimal]$MarginPercent / 100.0 
    if ($LowestCarrierCost -gt 0) {
        if ((1.0 - $marginDec) -ne 0) {
             try { $stdMarginPrice = [math]::Round(($LowestCarrierCost / (1.0 - $marginDec)), 2) }
             catch { Write-Warning "Error calculating standard margin price: $($_.Exception.Message)" }
        } else {
             Write-Warning "Cannot calculate standard margin price with 100% margin."
        }
    } else {
         Write-Verbose "Lowest carrier cost is zero or less, cannot calculate standard margin price."
    }
    $histAvgPrice = Get-HistoricalAveragePrice -OriginZip $OriginZip -DestinationZip $DestinationZip -Weight $Weight -FreightClass $FreightClass
    $finalPrice=$null; $reason="Error"
    if ($stdMarginPrice -ne $null -and $histAvgPrice -ne $null -and $histAvgPrice -gt 0) {
        if ($histAvgPrice -gt $stdMarginPrice) {
             $finalPrice = $histAvgPrice
             $reason = "Historical Avg Used (Higher than Std Margin Price)"
        } else {
             $finalPrice = $stdMarginPrice
             $reason = "Standard Margin ($($MarginPercent)%) Used (Higher/Equal to Hist Avg)"
        }
    } elseif ($stdMarginPrice -ne $null) {
        $finalPrice = $stdMarginPrice
        $reason = "Standard Margin ($($MarginPercent)%) Used (No Valid Historical Avg)"
    } elseif ($histAvgPrice -ne $null -and $histAvgPrice -gt 0) {
        $finalPrice = $histAvgPrice
        $reason = "Historical Avg Used (Std Margin Calculation Error)"
    } else {
        Write-Warning "Cannot determine a valid final price (Std Margin Price and Historical Avg are both invalid or zero)."
        $reason = "Calculation Error (No valid price source)"
    }
    $finalPriceRounded = if($finalPrice -ne $null){[math]::Round($finalPrice, 2)}else{$null}
    return [PSCustomObject]@{
        LowestCost = $LowestCarrierCost
        StandardMarginPrice = $stdMarginPrice
        HistoricalAvgPrice = $histAvgPrice
        MarginUsedPercent = $MarginPercent
        FinalPrice = $finalPriceRounded
        Reason = $reason
    }
}

function Write-QuoteToHistory {
     param(
        [Parameter(Mandatory)] [string]$Carrier,
        [Parameter(Mandatory)] [string]$TariffName,
        [Parameter(Mandatory)] [string]$OriginZip,
        [Parameter(Mandatory)] [string]$DestinationZip,
        [Parameter(Mandatory)] [double]$Weight,
        [Parameter(Mandatory)] [string]$FreightClass,
        [Parameter(Mandatory)] [decimal]$LowestCost, 
        [Parameter(Mandatory)] [decimal]$FinalQuotedPrice, 
        [Parameter(Mandatory)] [string]$QuoteTimestamp 
     )
    $logFileName = $Global:HistoricalQuotesLogFileName 
    if ([string]::IsNullOrWhiteSpace($logFileName)) { Write-Warning "HistoricalQuotesLogFileName not set in config. Cannot log quote."; return }
    $logPath = Join-Path $script:shipmentDataFolderPath $logFileName # Assumes $script:shipmentDataFolderPath is set

    if ([string]::IsNullOrWhiteSpace($OriginZip) -or $OriginZip.Length -lt 3 -or [string]::IsNullOrWhiteSpace($DestinationZip) -or $DestinationZip.Length -lt 3){ Write-Warning "Quote Log Skip: ZIP too short or missing."; return }
    if ($Weight -le 0) { Write-Warning "Quote Log Skip: Invalid Weight."; return }
    if ($LowestCost -le 0) { Write-Warning "Quote Log Skip: Invalid Lowest Cost."; return }
    if ($FinalQuotedPrice -le 0) { Write-Warning "Quote Log Skip: Invalid Final Quoted Price."; return }

    $oZip3=$OriginZip.Substring(0,3); $dZip3=$DestinationZip.Substring(0,3)
    $logEntry=[PSCustomObject]@{
        Timestamp = $QuoteTimestamp
        Carrier = $Carrier
        Tariff = $TariffName
        OriginZip3 = $oZip3
        DestZip3 = $dZip3
        Weight = $Weight
        FreightClass = $FreightClass
        LowestCost = $LowestCost 
        FinalQuotedPrice = $FinalQuotedPrice 
        OriginZipFull = $OriginZip
        DestinationZipFull = $DestinationZip
    }
    try {
        Ensure-DirectoryExists -Path (Split-Path $logPath -Parent)
        $fileExists = Test-Path $logPath -PathType Leaf
        if ($fileExists) {
             $logEntry | Export-Csv -Path $logPath -NoTypeInformation -Append -Encoding UTF8 -ErrorAction Stop
        } else {
             $logEntry | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        }
        Write-Verbose "Quote logged: $logPath"
    } catch {
        Write-Error "Quote log fail '$logPath': $($_.Exception.Message)"
    }
}

# --- Report Path Helper ---
function Get-ReportPath {
    param(
        [Parameter(Mandatory)][string]$BaseDir,
        [Parameter(Mandatory)][string]$Username,
        [Parameter(Mandatory)][string]$Carrier, 
        [Parameter(Mandatory)][string]$ReportType, 
        [string]$FilePrefix = $null, 
        [string]$FileExtension = "txt"
    )
    $userFolder = Join-Path -Path $BaseDir -ChildPath $Username
    try {
         Ensure-DirectoryExists -Path $userFolder 
    } catch {
         Write-Error "Failed to ensure user report directory '$userFolder' exists. Cannot generate report path."
         return $null 
    }
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $safeCarrier = $Carrier -replace '[^a-zA-Z0-9_-]', ''
    $safeReportType = $ReportType -replace '[^a-zA-Z0-9_-]', ''
    $safePrefix = if ($FilePrefix) { ($FilePrefix -replace '[^a-zA-Z0-9_-]', '').TrimStart('_').TrimEnd('_') + "_" } else { "" }
    $fileName = "{0}_{1}_{2}{3}.{4}" -f $safeCarrier, $safeReportType, $safePrefix, $timestamp, $FileExtension
    $fullPath = Join-Path -Path $userFolder -ChildPath $fileName
    Write-Verbose "Generated report path: $fullPath"
    return $fullPath
}

Write-Verbose "TMS General Helper Functions loaded."
