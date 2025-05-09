# TMS_Helpers_General.ps1
# Description: Contains general helper functions used across the TMS tool.

# --- Function to Load Key/Tariff Files from a Folder ---
Function Load-KeysFromFolder {
    param(
        [Parameter(Mandatory=$true)]
        [string]$KeysFolderPath,
        [Parameter(Mandatory=$true)]
        [string]$CarrierName # For logging/identification purposes
    )
    Write-Verbose "Loading $CarrierName keys from: $KeysFolderPath"
    $allKeyData = @{}
    if (-not (Test-Path $KeysFolderPath -PathType Container)) {
        Write-Warning "Key folder for $CarrierName not found at '$KeysFolderPath'."
        return $allKeyData # Return empty hashtable
    }

    $keyFiles = Get-ChildItem -Path $KeysFolderPath -Filter "*.txt" -File
    if ($keyFiles.Count -eq 0) {
        Write-Warning "No .txt key files found in '$KeysFolderPath' for $CarrierName."
    }

    foreach ($file in $keyFiles) {
        $keyName = $file.BaseName # Use filename without extension as a default key name
        $keyData = @{
            TariffFileName = $file.Name # Store the original filename
        }
        $validPairsFound = $false
        $fileContent = Get-Content -Path $file.FullName -ErrorAction SilentlyContinue

        if ($null -eq $fileContent) {
            Write-Warning "Could not read content from '$($file.Name)' or file is empty."
            continue # Skip to the next file
        }

        foreach ($line in $fileContent) {
            $trimmedLine = $line.Trim()
            if ($trimmedLine -eq "" -or $trimmedLine.StartsWith("#")) {
                continue # Skip empty lines and comments
            }

            # Try to split by ':' or '=' , limit to 2 parts to handle values that might contain the delimiter
            $parts = $trimmedLine -split '[:=]', 2
            if ($parts.Count -eq 2) {
                $fileKey = $parts[0].Trim()
                $fileValue = $parts[1].Trim()
                if (-not [string]::IsNullOrWhiteSpace($fileKey)) {
                    $keyData[$fileKey] = $fileValue
                    if ($fileKey -eq "Name" -and -not [string]::IsNullOrWhiteSpace($fileValue)) {
                        $keyName = $fileValue # Use "Name" field from file as the primary key if present
                    }
                    $validPairsFound = $true
                } else {
                    Write-Warning "Skipping line (empty key after trim) in '$($file.Name)': '$line'"
                }
            } else {
                Write-Warning "Skipping line (no valid Key=Value or Key:Value pair) in '$($file.Name)': '$line'"
            }
        }

        if ($validPairsFound) {
            if ($allKeyData.ContainsKey($keyName)) {
                Write-Warning "Duplicate key name '$keyName' (from file '$($file.Name)' or 'Name' field). Overwriting previous entry. Ensure unique names or filenames."
            }
            $allKeyData[$keyName] = $keyData
            Write-Verbose "Loaded key '$keyName' from '$($file.Name)' for $CarrierName."
        } else {
            Write-Warning "No valid Key=Value pairs found or parsed in '$($file.Name)'."
        }
    }
    Write-Verbose "Finished loading $CarrierName keys. Total loaded: $($allKeyData.Count)"
    return $allKeyData
}

Function Ensure-DirectoryExists {
    param([string]$Path)
    if (-not (Test-Path -Path $Path -PathType Container)) {
        try {
            New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Verbose "Created directory: $Path"
        } catch {
            throw "Failed to create directory '$Path'. Error: $($_.Exception.Message)"
        }
    }
}

Function Get-PermittedKeys {
    param(
        [Parameter(Mandatory=$true)][hashtable]$AllKeys, # e.g., $script:allCentralKeys
        [Parameter(Mandatory=$true)][array]$AllowedKeyNames # e.g., $customerProfile.AllowedCentralKeys
    )
    $permitted = @{}
    if ($null -eq $AllowedKeyNames -or $AllowedKeyNames.Count -eq 0) {
        return $permitted # No keys are allowed
    }
    foreach ($allowedName in $AllowedKeyNames) {
        if ($AllKeys.ContainsKey($allowedName)) {
            $permitted[$allowedName] = $AllKeys[$allowedName]
        } else {
            # Attempt to match by TariffFileName if direct name match fails
            # This handles cases where AllowedKeyNames might store the filename
            $foundByFileName = $false
            foreach($keyEntry in $AllKeys.GetEnumerator()){
                if($keyEntry.Value -is [hashtable] -and $keyEntry.Value.ContainsKey('TariffFileName') -and $keyEntry.Value.TariffFileName -eq $allowedName){
                    $permitted[$keyEntry.Key] = $keyEntry.Value # Add using the actual key name
                    $foundByFileName = $true
                    break
                }
            }
            if(-not $foundByFileName){
                 Write-Warning "Allowed key/tariff '$allowedName' not found in the loaded set of all keys."
            }
        }
    }
    return $permitted
}

Function Select-CsvFile {
    param (
        [string]$DialogTitle = "Select CSV File",
        [string]$InitialDirectory = $PSScriptRoot # $PSScriptRoot might be null if run directly in console, handle accordingly
    )
    # Ensure $InitialDirectory has a fallback if $PSScriptRoot is null
    if ([string]::IsNullOrWhiteSpace($InitialDirectory)) {
        try {
            $InitialDirectory = Get-Location # Default to current working directory
        } catch {
            $InitialDirectory = "." # Absolute fallback
            Write-Warning "Select-CsvFile: Could not Get-Location, defaulting InitialDirectory to '.'"
        }
    }
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Title = $DialogTitle
        $openFileDialog.InitialDirectory = $InitialDirectory
        $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            return $openFileDialog.FileName
        }
    } catch {
        Write-Warning "Select-CsvFile: Error displaying OpenFileDialog. Ensure GUI components are available. Error: $($_.Exception.Message)"
    }
    return $null
}

Function Open-FileExplorer {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    if (Test-Path $Path) {
        try {
            Invoke-Item $Path
        } catch {
            Write-Warning "Open-FileExplorer: Could not open path '$Path'. Error: $($_.Exception.Message)"
        }
    } else {
        Write-Warning "Path not found: $Path"
    }
}

function Get-MinimumRate {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$RateResults
    )
    Write-Verbose "Get-MinimumRate called with $($RateResults.Count) rate results."
    if ($RateResults.Count -eq 0) {
        Write-Warning "Get-MinimumRate: No rate results provided."
        return $null
    }

    $lowestCost = $null
    $selectedTariffName = $null

    foreach ($tariffName in $RateResults.Keys) {
        $currentCost = $RateResults[$tariffName]
        if ($currentCost -is [decimal] -or $currentCost -is [double] -or $currentCost -is [int]) {
            if ($null -eq $lowestCost -or [decimal]$currentCost -lt [decimal]$lowestCost) {
                $lowestCost = [decimal]$currentCost
                $selectedTariffName = $tariffName
            }
        } else {
            Write-Warning "Get-MinimumRate: Invalid cost value for tariff '$tariffName': $currentCost"
        }
    }

    if ($null -ne $selectedTariffName) {
        Write-Verbose "Get-MinimumRate: Lowest cost is $lowestCost from tariff '$selectedTariffName'."
        return [PSCustomObject]@{
            TariffName = $selectedTariffName
            Cost       = $lowestCost
        }
    } else {
        Write-Warning "Get-MinimumRate: Could not determine a minimum rate from the provided results."
        return $null
    }
} 

function Calculate-QuotePrice {
    param(
        [Parameter(Mandatory=$true)][decimal]$LowestCarrierCost,
        [Parameter(Mandatory=$true)][string]$OriginZip,
        [Parameter(Mandatory=$true)][string]$DestinationZip,
        [Parameter(Mandatory=$true)][decimal]$Weight,
        [Parameter(Mandatory=$true)][string]$FreightClass,
        [Parameter(Mandatory=$true)][double]$MarginPercent,
        [decimal]$MinimumProfit = $script:DefaultMinProfit 
    )
    Write-Verbose "Calculate-QuotePrice called. Cost: $LowestCarrierCost, Margin: $MarginPercent%"

    if ($LowestCarrierCost -le 0) {
        Write-Warning "Calculate-QuotePrice: LowestCarrierCost is zero or negative. Cannot calculate price."
        return [PSCustomObject]@{ FinalPrice = $null; CalculationError = "Invalid base cost." }
    }
    
    $MarginPercentAsDouble = 0.0
    try { 
        $MarginPercentAsDouble = [System.Convert]::ToDouble($MarginPercent) 
    } catch { 
        Write-Warning "Calculate-QuotePrice: Invalid margin format for value '$($MarginPercent)'. Error: $($_.Exception.Message). Defaulting margin to 0."
        $MarginPercentAsDouble = 0.0
    }

    if ($MarginPercentAsDouble -lt 0 -or $MarginPercentAsDouble -ge 100) {
        Write-Warning "Calculate-QuotePrice: MarginPercent '$MarginPercentAsDouble' is out of valid range (0-99.9). Defaulting to 0% for safety."
        $MarginPercentAsDouble = 0.0
    }

    $marginDecimal = $MarginPercentAsDouble / 100.0
    $calculatedPrice = $null

    if ([System.Math]::Abs(1.0 - $marginDecimal) -lt 0.00001) { 
        Write-Warning "Calculate-QuotePrice: Cannot calculate price with 100% margin (division by zero)."
        return [PSCustomObject]@{ FinalPrice = $null; CalculationError = "100% margin results in division by zero." }
    }
    
    $calculatedPrice = $LowestCarrierCost / (1.0 - $marginDecimal)

    $effectiveMinProfit = $MinimumProfit
    if ($null -eq $effectiveMinProfit) {
        if ((Get-Variable -Name "script:DefaultMinProfit" -ErrorAction SilentlyContinue) -and ($null -ne $script:DefaultMinProfit)) {
             try { $effectiveMinProfit = [decimal]$script:DefaultMinProfit } catch { $effectiveMinProfit = 50.0; Write-Warning "Calculate-QuotePrice: Error converting script:DefaultMinProfit. Using 50."}
        } else {
            Write-Warning "Calculate-QuotePrice: script:DefaultMinProfit not found or null. Using default of 50."
            $effectiveMinProfit = 50.0 
        }
    }

    $profit = $calculatedPrice - $LowestCarrierCost
    if ($profit -lt $effectiveMinProfit) {
        $calculatedPrice = $LowestCarrierCost + $effectiveMinProfit
        Write-Verbose "Calculate-QuotePrice: Price adjusted to meet minimum profit of $effectiveMinProfit. New price: $calculatedPrice"
    }

    $finalPriceRounded = [Math]::Round($calculatedPrice, 2)
    Write-Verbose "Calculate-QuotePrice: Final Quoted Price: $finalPriceRounded"
    return [PSCustomObject]@{
        FinalPrice       = $finalPriceRounded
        BaseCost         = $LowestCarrierCost
        AppliedMargin    = $MarginPercentAsDouble 
        CalculatedProfit = [Math]::Round(($finalPriceRounded - $LowestCarrierCost), 2)
    }
} 

function Write-QuoteToHistory {
    param(
        [Parameter(Mandatory=$true)][string]$Carrier,
        [Parameter(Mandatory=$true)][string]$Tariff,
        [Parameter(Mandatory=$true)][string]$OriginZip,       
        [Parameter(Mandatory=$true)][string]$DestinationZip,  
        [Parameter(Mandatory=$true)][decimal]$Weight,
        [Parameter(Mandatory=$true)][string]$FreightClass,
        [Parameter(Mandatory=$true)][decimal]$LowestCost,
        [Parameter(Mandatory=$true)][decimal]$FinalQuotedPrice,
        [Parameter(Mandatory=$true)][string]$QuoteTimestamp,  
        [string]$OriginZipFull = $null,      
        [string]$DestinationZipFull = $null  
    )

    Write-Verbose "Write-QuoteToHistory called:"
    Write-Verbose "  Timestamp: $QuoteTimestamp, Carrier: $Carrier, Tariff: $Tariff"
    Write-Verbose "  OriginZip3: $OriginZip, DestZip3: $DestinationZip, Weight: $Weight, Class: $FreightClass"
    Write-Verbose "  LowestCost: $LowestCost, FinalQuotedPrice: $FinalQuotedPrice"
    Write-Verbose "  OriginZipFull: $($OriginZipFull | Out-String), DestinationZipFull: $($DestinationZipFull | Out-String)"

    $historyFilePath = $null
    try {
        $basePath = $PSScriptRoot
        if ([string]::IsNullOrWhiteSpace($basePath) -and $MyInvocation.MyCommand.Path) {
            $basePath = Split-Path $MyInvocation.MyCommand.Path -Parent
        }
        if ([string]::IsNullOrWhiteSpace($basePath)) {
            Write-Warning "Write-QuoteToHistory: Could not determine script base path. Using current directory as fallback."
            $basePath = "." 
        }

        $shipmentDataFolderName = "shipmentData" 
        $historicalQuotesFileName = "historical_quotes_generated.csv" 

        if (Get-Variable -Name "script:defaultShipmentDataFolderName" -Scope Script -ErrorAction SilentlyContinue) {
            $shipmentDataFolderName = $script:defaultShipmentDataFolderName
        } elseif (Get-Variable -Name "defaultShipmentDataFolderName" -Scope Global -ErrorAction SilentlyContinue) {
            $shipmentDataFolderName = $global:defaultShipmentDataFolderName
        }
        
        if (Get-Variable -Name "script:defaultHistoricalQuotesFileName" -Scope Script -ErrorAction SilentlyContinue) {
            $historicalQuotesFileName = $script:defaultHistoricalQuotesFileName
        } elseif (Get-Variable -Name "defaultHistoricalQuotesFileName" -Scope Global -ErrorAction SilentlyContinue) {
            $historicalQuotesFileName = $global:defaultHistoricalQuotesFileName
        }

        $fullShipmentDataPath = Join-Path -Path $basePath -ChildPath $shipmentDataFolderName
        $historyFilePath = Join-Path -Path $fullShipmentDataPath -ChildPath $historicalQuotesFileName

        Ensure-DirectoryExists -Path $fullShipmentDataPath

        $newRecord = [PSCustomObject]@{
            "Timestamp"          = $QuoteTimestamp
            "Carrier"            = $Carrier
            "Tariff"             = $Tariff
            "OriginZip3"         = $OriginZip
            "DestZip3"           = $DestinationZip
            "Weight"             = $Weight
            "FreightClass"       = $FreightClass
            "LowestCost"         = [Math]::Round($LowestCost, 2)
            "FinalQuotedPrice"   = [Math]::Round($FinalQuotedPrice, 2)
            "OriginZipFull"      = if ([string]::IsNullOrWhiteSpace($OriginZipFull)) { $OriginZip } else { $OriginZipFull }
            "DestinationZipFull" = if ([string]::IsNullOrWhiteSpace($DestinationZipFull)) { $DestinationZip } else { $DestinationZipFull }
        }

        if (-not (Test-Path $historyFilePath)) {
            $newRecord | Export-Csv -Path $historyFilePath -NoTypeInformation -Encoding UTF8
            Write-Verbose "Created historical quotes file and wrote record: $historyFilePath"
        } else {
            $newRecord | Export-Csv -Path $historyFilePath -NoTypeInformation -Append -Encoding UTF8
            Write-Verbose "Appended quote to historical file: $historyFilePath"
        }
    } catch {
        Write-Warning "Write-QuoteToHistory: Error writing to historical quotes file '$($historyFilePath | Out-String)'. Details: $($_.Exception.Message)"
    }
} 

# --- ADDED Get-ReportPath FUNCTION ---
Function Get-ReportPath {
    param(
        [Parameter(Mandatory=$true)]
        [string]$BaseDir,
        [Parameter(Mandatory=$true)]
        [string]$Username,
        [Parameter(Mandatory=$true)]
        [string]$Carrier,
        [Parameter(Mandatory=$true)]
        [string]$ReportType,
        [string]$FilePrefix = ""
    )
    try {
        # Sanitize inputs for directory/file names
        $safeUsername = $Username -replace '[^a-zA-Z0-9_-]', ''
        $safeCarrier = $Carrier -replace '[^a-zA-Z0-9_-]', ''
        $safeReportType = $ReportType -replace '[^a-zA-Z0-9_-]', ''
        $safeFilePrefix = $FilePrefix -replace '[^a-zA-Z0-9_-]', ''

        # Construct directory path: BaseDir/Username/Carrier
        $reportDirectory = Join-Path -Path $BaseDir -ChildPath $safeUsername
        $reportDirectory = Join-Path -Path $reportDirectory -ChildPath $safeCarrier
        
        Ensure-DirectoryExists -Path $reportDirectory # Ensure the directory exists

        # Construct filename
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $fileName = "${timestamp}_${safeUsername}_${safeCarrier}_${safeReportType}"
        if (-not [string]::IsNullOrWhiteSpace($safeFilePrefix)) {
            $fileName += "_${safeFilePrefix}"
        }
        $fileName += ".txt"

        return Join-Path -Path $reportDirectory -ChildPath $fileName
    } catch {
        Write-Warning "Get-ReportPath: Error generating report path. Details: $($_.Exception.Message)"
        return $null
    }
}

Write-Verbose "TMS General Helper Functions loaded."
