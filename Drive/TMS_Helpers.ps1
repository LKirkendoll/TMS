# TMS_Helpers.ps1
# Description: Contains reusable helper functions for the TMS Tool.
#              Includes API calls, data normalization, key file parsing, etc.
#              This file should be dot-sourced by the main entry script (TMS_GUI.ps1).

# --- .NET Assembly Loading ---
# Ensure necessary assemblies are loaded for UI elements and other functionalities.
try {
    Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
    Add-Type -AssemblyName System.Drawing # For GUI elements if any are used directly here
} catch {
    Write-Error "FATAL ERROR in TMS_Helpers.ps1: Failed to load required .NET Assemblies (System.Windows.Forms, System.Drawing). Ensure .NET Framework is available. Error: $($_.Exception.Message)"
    # This is a critical failure; the main script should handle script termination.
    throw "Assembly load failed in TMS_Helpers.ps1."
}

# --- File/Folder Functions ---

function Ensure-DirectoryExists {
    # Creates a directory if it doesn't exist.
    param( [Parameter(Mandatory=$true)][string]$Path )
    if (-not (Test-Path -Path $Path -PathType Container)) {
        Write-Warning "Required folder '$(Split-Path -Path $Path -Leaf)' ('$Path') not found. Creating it..."
        try {
            New-Item -Path $Path -ItemType Directory -Force -ErrorAction Stop | Out-Null
            Write-Verbose "Successfully created directory: $Path"
        } catch {
            Write-Error "Failed to create folder '$Path': $($_.Exception.Message)"
            throw "Directory creation failed for $Path." # Allow calling script to handle termination
        }
    }
}

# --- User Interface Helper Functions ---

function Select-CsvFile {
    # Prompts the user to select a CSV file using a standard Windows Forms OpenFileDialog.
    param(
        [string]$DialogTitle = "Select CSV File",
        [string]$InitialDirectory # Should be passed from calling script, e.g., $script:shipmentDataFolderPath
    )

    $actualInitialDirectory = $InitialDirectory
    if ([string]::IsNullOrWhiteSpace($actualInitialDirectory) -or -not (Test-Path $actualInitialDirectory -PathType Container)) {
         Write-Warning "Select-CsvFile: Initial directory '$actualInitialDirectory' is invalid or not provided. Defaulting to script root: '$script:scriptRoot'."
         $actualInitialDirectory = $script:scriptRoot # $script:scriptRoot should be set by the main entry script
         if ([string]::IsNullOrWhiteSpace($actualInitialDirectory) -or -not (Test-Path $actualInitialDirectory -PathType Container)){
            Write-Warning "Select-CsvFile: Script root '$($script:scriptRoot)' is also invalid. Defaulting to C:\."
            $actualInitialDirectory = "C:\"
         }
    }

    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = $DialogTitle
    $dialog.InitialDirectory = $actualInitialDirectory
    $dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $dialog.FileName
    } else {
        Write-Warning "File selection cancelled by user."
        return $null
    }
}

function Clear-HostAndDrawHeader {
    # Clears the host screen and draws a standard header (for console mode).
    param(
        [Parameter(Mandatory=$true)] [string]$Title,
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

function Write-LoadingBar {
    # Displays a simple text-based loading bar using Write-Progress (for console mode).
    param(
        [Parameter(Mandatory=$true)][int]$PercentComplete,
        [string]$Message = "Processing..."
    )
    $validPercent = $PercentComplete
    if ($validPercent -lt 0) { $validPercent = 0 }
    if ($validPercent -gt 100) { $validPercent = 100 }
    Write-Progress -Activity $Message -Status "$validPercent% Complete" -PercentComplete $validPercent
}

function Open-FileExplorer {
    # Opens Windows File Explorer to the specified path.
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

# --- Data Handling: Key/Tariff Loading ---

function Load-KeysFromFolder {
    # Loads all .txt key files from a specified folder, parsing them into hashtables.
    param(
        [Parameter(Mandatory=$true)][string]$KeysFolderPath,
        [Parameter(Mandatory=$true)][string]$CarrierName
    )
    $loadedKeysAndMargins = @{}
    Write-Verbose "Loading keys/margins for $CarrierName from: $KeysFolderPath"
    if (-not (Test-Path -Path $KeysFolderPath -PathType Container)) {
        Write-Warning "Key folder '$KeysFolderPath' for $CarrierName not found."
        return $loadedKeysAndMargins # Return empty
    }

    $keyFiles = Get-ChildItem -Path $KeysFolderPath -Filter "*.txt" -File -ErrorAction SilentlyContinue
    if ($keyFiles) {
        foreach ($file in $keyFiles) {
            $keyNameFromFile = $file.BaseName
            $keyDataHashtable = @{}
            $isFirstLine = $true # To handle potential BOM or special characters in the first key name
            try {
                $lines = Get-Content -Path $file.FullName -ErrorAction Stop
                foreach ($line in $lines) {
                    $trimmedLine = $line.Trim()
                    if ([string]::IsNullOrWhiteSpace($trimmedLine) -or $trimmedLine.StartsWith('#')) { continue } # Skip comments/empty

                    $equalsIndex = $trimmedLine.IndexOf('=')
                    if ($equalsIndex -gt 0) {
                        $key = $trimmedLine.Substring(0, $equalsIndex).Trim()
                        $value = $trimmedLine.Substring($equalsIndex + 1).Trim()

                        if ($isFirstLine -and $key -match '^[^a-zA-Z0-9]*(.*)') { # Attempt to clean non-alphanumeric prefixes from first key
                            $cleanedKey = $Matches[1].Trim()
                            if (-not [string]::IsNullOrEmpty($cleanedKey)) { $key = $cleanedKey }
                        }
                        $isFirstLine = $false

                        if (-not [string]::IsNullOrEmpty($key)) { $keyDataHashtable[$key] = $value }
                    } else { Write-Warning "Skipping malformed line (no '=') in key file '$($file.Name)': $line" }
                }

                if ($keyDataHashtable.Count -gt 0) {
                    $keyDataHashtable['TariffFileName'] = $keyNameFromFile # Store original filename for reference
                    if (-not $keyDataHashtable.ContainsKey('Name')) { $keyDataHashtable['Name'] = $keyNameFromFile } # Ensure 'Name' property

                    # Validate MarginPercent
                    if (-not $keyDataHashtable.ContainsKey('MarginPercent')) {
                        Write-Warning "Key file '$($file.Name)' is missing 'MarginPercent'. Default margin will be used if applicable."
                    } elseif (($keyDataHashtable.MarginPercent -as [double]) -eq $null) {
                        Write-Warning "Invalid (non-numeric) 'MarginPercent' value found in '$($file.Name)'."
                    }

                    # Carrier-specific key validation
                    switch ($CarrierName) {
                        "Central Transport" {
                            if (-not $keyDataHashtable.ContainsKey('accessCode')) { Write-Warning "Central key file '$($file.Name)' missing 'accessCode'."}
                            if (-not $keyDataHashtable.ContainsKey('customerNumber')) { Write-Warning "Central key file '$($file.Name)' missing 'customerNumber'."}
                        }
                        "SAIA" {
                             if (-not $keyDataHashtable.ContainsKey('UserID') -and -not $keyDataHashtable.ContainsKey('RQKey')) { Write-Warning "SAIA key file '$($file.Name)' missing 'UserID' and 'RQKey'. One is required."}
                             if ($keyDataHashtable.ContainsKey('UserID') -and -not $keyDataHashtable.ContainsKey('Password')) { Write-Warning "SAIA key file '$($file.Name)' has 'UserID' but is missing 'Password'."}
                        }
                        "RL Carriers" {
                            if (-not $keyDataHashtable.ContainsKey('APIKey')) { Write-Warning "R+L Carriers key file '$($file.Name)' missing 'APIKey'."}
                        }
                        "Averitt" {
                            if (-not $keyDataHashtable.ContainsKey('APIKey')) { Write-Warning "Averitt key file '$($file.Name)' missing 'APIKey'." }
                        }
                    }
                    $loadedKeysAndMargins[$keyNameFromFile] = $keyDataHashtable
                    Write-Verbose "Successfully loaded key data for '$keyNameFromFile' ($CarrierName)."
                } else { Write-Warning "No valid Key=Value pairs found in key file '$($file.Name)'." }
            } catch { Write-Warning "Could not process key file '$($file.Name)': $($_.Exception.Message)" }
        }
    } else { Write-Verbose "No .txt key files found in '$KeysFolderPath' for $CarrierName." }
    Write-Host "Loaded $($loadedKeysAndMargins.Count) $CarrierName key(s)/account(s)." -ForegroundColor Gray
    return $loadedKeysAndMargins
}

function Get-PermittedKeys {
    # Filters a list of all carrier keys against a list of allowed key names for a user/customer.
    param(
        [Parameter(Mandatory=$true)][hashtable]$AllKeys,
        [Parameter(Mandatory=$true)][array]$AllowedKeyNames
    )
    $permittedKeys = @{}
    if ($null -ne $AllowedKeyNames) {
        foreach ($allowedName in $AllowedKeyNames) {
             if ([string]::IsNullOrWhiteSpace($allowedName)) { continue } # Skip empty names
             if ($AllKeys.ContainsKey($allowedName)) {
                 if ($AllKeys[$allowedName] -is [hashtable]) {
                     $permittedKeys[$allowedName] = $AllKeys[$allowedName]
                 } else { Write-Warning "Data for key '$allowedName' was not in the expected hashtable format." }
             } else { Write-Warning "User/Customer has permission for key '$allowedName', but this key was not found in the loaded keys for the carrier." }
        }
    }
    return $permittedKeys
}

# --- Data Normalization Functions (Carrier Specific) ---
# These functions prepare data from general CSVs for specific carrier API calls.

function Load-And-Normalize-CentralData {
    param([Parameter(Mandatory=$true)][string]$CsvPath)
    Write-Host "`nLoading Central data from: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan
    $requiredColumns = @("Origin Postal Code", "Destination Postal Code", "Total Weight", "Freight Class 1")
    try {
        $rawData = Import-Csv -Path $CsvPath -ErrorAction Stop
        Write-Host " -> Rows found: $($rawData.Count)." -ForegroundColor Gray
        if ($rawData.Count -eq 0) { Write-Warning "CSV file is empty."; return @() }

        $headers = $rawData[0].PSObject.Properties.Name
        $missingColumns = $requiredColumns | Where-Object { $_ -notin $headers }
        if ($missingColumns) {
            Write-Error "CSV file '$CsvPath' for Central is missing required columns: $($missingColumns -join ', ')"
            return $null
        }

        $normalizedData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing Central data..." -ForegroundColor Gray
        $invalidRowCount = 0
        for ($i = 0; $i -lt $rawData.Count; $i++) {
            $row = $rawData[$i]
            $rowNumber = $i + 1
            $originZip = ($row."Origin Postal Code" -replace '\s','').Trim()
            $destZip = ($row."Destination Postal Code" -replace '\s','').Trim()
            $weightString = ($row."Total Weight" -replace '\s','').Trim()
            $classString = ($row."Freight Class 1" -replace '\s','').Trim()
            $weightDecimal = $null

            $isValidRow = $true
            if (-not ($originZip -match '^\d{5}')) { Write-Verbose "Skipping Central Row ${rowNumber}: Invalid Origin ZIP '$($row."Origin Postal Code")'"; $isValidRow = $false }
            if ($isValidRow -and -not ($destZip -match '^\d{5}')) { Write-Verbose "Skipping Central Row ${rowNumber}: Invalid Destination ZIP '$($row."Destination Postal Code")'"; $isValidRow = $false }
            if ($isValidRow -and [string]::IsNullOrWhiteSpace($classString)) { Write-Verbose "Skipping Central Row ${rowNumber}: Missing Freight Class"; $isValidRow = $false }
            if ($isValidRow) {
                try { $weightDecimal = [decimal]$weightString; if($weightDecimal -le 0) { throw "Weight must be positive." } }
                catch { Write-Verbose "Skipping Central Row ${rowNumber}: Invalid Total Weight '$weightString'. Error: $($_.Exception.Message)"; $isValidRow = $false }
            }

            if ($isValidRow) {
                 $normalizedData.Add([PSCustomObject]@{
                    "Origin Postal Code"    = $originZip       # Keep original names if CTX API uses them
                    "Destination Postal Code" = $destZip
                    "Total Weight"          = $weightDecimal   # Store numeric weight
                    "Freight Class 1"       = $classString
                 })
            } else { $invalidRowCount++ }
        }
        if ($invalidRowCount -gt 0) { Write-Warning " -> Skipped $invalidRowCount Central rows due to missing/invalid essential data." }
        Write-Host " -> OK: $($normalizedData.Count) Central rows normalized." -ForegroundColor Green
        return $normalizedData
    } catch { Write-Error "Error processing Central CSV '$CsvPath': $($_.Exception.Message)"; return $null }
}

function Load-And-Normalize-SAIAData {
    param([Parameter(Mandatory=$true)][string]$CsvPath)
    Write-Host "`nLoading SAIA data from: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan
    $requiredColumns = @("Origin Postal Code", "Destination Postal Code", "Total Weight", "Freight Class 1", "Origin City", "Origin State", "Destination City", "Destination State")
    $optionalColumns = @("Total Units", "Total Density") # For dimensional calculations
    try {
        $rawData = Import-Csv -Path $CsvPath -ErrorAction Stop
        Write-Host " -> Rows found: $($rawData.Count)." -ForegroundColor Gray
        if ($rawData.Count -eq 0) { Write-Warning "CSV file is empty."; return @() }

        $headers = $rawData[0].PSObject.Properties.Name
        $missingRequired = $requiredColumns | Where-Object { $_ -notin $headers }
        if ($missingRequired) { Write-Error "CSV file '$CsvPath' for SAIA is missing required columns: $($missingRequired -join ', ')"; return $null }
        $missingOptional = $optionalColumns | Where-Object { $_ -notin $headers }
        if ($missingOptional) { Write-Warning "SAIA CSV '$CsvPath' missing optional columns for dimension calculation: $($missingOptional -join ', ')" }

        $normalizedData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing SAIA data..." -ForegroundColor Gray
        $invalidRowCount = 0
        for ($i = 0; $i -lt $rawData.Count; $i++) {
            $row = $rawData[$i]; $rowNumber = $i + 1
            $originZip = ($row."Origin Postal Code" -replace '\s','').Trim(); $destZip = ($row."Destination Postal Code" -replace '\s','').Trim()
            $weightString = ($row."Total Weight" -replace '\s','').Trim(); $classString = ($row."Freight Class 1" -replace '\s','').Trim()
            $originCity = $row."Origin City".Trim(); $originState = $row."Origin State".Trim().ToUpper()
            $destCity = $row."Destination City".Trim(); $destState = $row."Destination State".Trim().ToUpper()
            $weightDecimal = $null; $unitsInt = $null; $densityDouble = $null; $isValidRow = $true

            if (-not ($originZip -match '^\d{5}')) { Write-Verbose "Skip SAIA Row ${rowNumber}: Invalid Origin ZIP"; $isValidRow = $false }
            if ($isValidRow -and -not ($destZip -match '^\d{5}')) { Write-Verbose "Skip SAIA Row ${rowNumber}: Invalid Dest ZIP"; $isValidRow = $false }
            if ($isValidRow -and [string]::IsNullOrWhiteSpace($classString)) { Write-Verbose "Skip SAIA Row ${rowNumber}: Missing Class"; $isValidRow = $false }
            if ($isValidRow -and [string]::IsNullOrWhiteSpace($originCity)) { Write-Verbose "Skip SAIA Row ${rowNumber}: Missing Origin City"; $isValidRow = $false }
            if ($isValidRow -and -not ($originState -match '^[A-Z]{2}$')) { Write-Verbose "Skip SAIA Row ${rowNumber}: Invalid Origin State"; $isValidRow = $false }
            if ($isValidRow -and [string]::IsNullOrWhiteSpace($destCity)) { Write-Verbose "Skip SAIA Row ${rowNumber}: Missing Dest City"; $isValidRow = $false }
            if ($isValidRow -and -not ($destState -match '^[A-Z]{2}$')) { Write-Verbose "Skip SAIA Row ${rowNumber}: Invalid Dest State"; $isValidRow = $false }
            if ($isValidRow) { try { $weightDecimal = [decimal]$weightString; if($weightDecimal -le 0){throw "Wt <=0"} } catch { Write-Verbose "Skip SAIA Row ${rowNumber}: Invalid Weight"; $isValidRow = $false } }

            if (-not $isValidRow) { $invalidRowCount++; continue }

            # Optional fields for SAIA details array
            if ($headers -contains 'Total Units') { try { $unitsInt = [int]($row.'Total Units'.Trim()); if ($unitsInt -le 0) {$unitsInt = 1} } catch { $unitsInt = 1} } else { $unitsInt = 1 }
            # SAIA API expects length, width, height for its 'details' array. If not in CSV, use defaults.
            $itemLength = 1.0; $itemWidth = 1.0; $itemHeight = 1.0; # Defaults

            $detailItem = @{
                weight = [int]$weightDecimal # SAIA API often expects integer weight in details
                class  = $classString       # SAIA API expects class as string in details
                length = $itemLength
                width  = $itemWidth
                height = $itemHeight
                units  = $unitsInt
            }

            $normalizedData.Add([PSCustomObject]@{
                OriginZip        = $originZip; DestinationZip   = $destZip
                OriginCity       = $originCity; OriginState      = $originState
                DestinationCity  = $destCity; DestinationState = $destState
                details          = @($detailItem) # SAIA API expects an array of detail items
                'Total Weight'   = $weightDecimal # For reference
                'Freight Class 1'= $classString   # For reference
            })
        }
        if ($invalidRowCount -gt 0) { Write-Warning " -> Skipped $invalidRowCount SAIA rows due to missing/invalid essential data." }
        Write-Host " -> OK: $($normalizedData.Count) SAIA rows normalized." -ForegroundColor Green
        return $normalizedData
    } catch { Write-Error "Error processing SAIA CSV '$CsvPath': $($_.Exception.Message)"; return $null }
}

function Load-And-Normalize-RLData {
    param([Parameter(Mandatory=$true)][string]$CsvPath)
    Write-Host "`nLoading R+L data from: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan
    # R+L requires City/State in addition to Zips for its API.
    $requiredColumns = @("Origin Postal Code", "Destination Postal Code", "Total Weight", "Freight Class 1", "Origin City", "Origin State", "Destination City", "Destination State")
    try {
        $rawData = Import-Csv -Path $CsvPath -ErrorAction Stop
        Write-Host " -> Rows found: $($rawData.Count)." -ForegroundColor Gray
        if ($rawData.Count -eq 0) { Write-Warning "CSV file is empty."; return @() }

        $headers = $rawData[0].PSObject.Properties.Name
        $missingRequired = $requiredColumns | Where-Object { $_ -notin $headers }
        if ($missingRequired) { Write-Error "CSV file '$CsvPath' for R+L is missing required columns: $($missingRequired -join ', ')"; return $null }

        $normalizedData = [System.Collections.Generic.List[object]]::new()
        Write-Host " -> Normalizing R+L data..." -ForegroundColor Gray
        $invalidRowCount = 0
        for ($i = 0; $i -lt $rawData.Count; $i++) {
            $row = $rawData[$i]; $rowNumber = $i + 1
            $originZip = ($row."Origin Postal Code" -replace '\s','').Trim(); $destZip = ($row."Destination Postal Code" -replace '\s','').Trim()
            $weightString = ($row."Total Weight" -replace '\s','').Trim(); $classString = ($row."Freight Class 1" -replace '\s','').Trim()
            $originCity = $row."Origin City".Trim(); $originState = $row."Origin State".Trim().ToUpper()
            $destCity = $row."Destination City".Trim(); $destState = $row."Destination State".Trim().ToUpper()
            $weightDecimal = $null; $isValidRow = $true

            if (-not ($originZip -match '^\d{5}')) { Write-Verbose "Skip R+L Row ${rowNumber}: Invalid Origin ZIP"; $isValidRow = $false }
            if ($isValidRow -and -not ($destZip -match '^\d{5}')) { Write-Verbose "Skip R+L Row ${rowNumber}: Invalid Dest ZIP"; $isValidRow = $false }
            if ($isValidRow -and [string]::IsNullOrWhiteSpace($classString)) { Write-Verbose "Skip R+L Row ${rowNumber}: Missing Class"; $isValidRow = $false }
            if ($isValidRow -and [string]::IsNullOrWhiteSpace($originCity)) { Write-Verbose "Skip R+L Row ${rowNumber}: Missing Origin City"; $isValidRow = $false }
            if ($isValidRow -and -not ($originState -match '^[A-Z]{2}$')) { Write-Verbose "Skip R+L Row ${rowNumber}: Invalid Origin State"; $isValidRow = $false }
            if ($isValidRow -and [string]::IsNullOrWhiteSpace($destCity)) { Write-Verbose "Skip R+L Row ${rowNumber}: Missing Dest City"; $isValidRow = $false }
            if ($isValidRow -and -not ($destState -match '^[A-Z]{2}$')) { Write-Verbose "Skip R+L Row ${rowNumber}: Invalid Dest State"; $isValidRow = $false }
            if ($isValidRow) { try { $weightDecimal = [decimal]$weightString; if($weightDecimal -le 0){throw "Wt <=0"} } catch { Write-Verbose "Skip R+L Row ${rowNumber}: Invalid Weight"; $isValidRow = $false } }

            if (-not $isValidRow) { $invalidRowCount++; continue }

            # R+L API takes optional fields directly in its ShipmentDetails parameter for Invoke-RLApi
            $entry = [PSCustomObject]@{
                OriginZip        = $originZip; DestinationZip   = $destZip
                Weight           = $weightDecimal; Class            = $classString
                OriginCity       = $originCity; OriginState      = $originState
                DestinationCity  = $destCity; DestinationState = $destState
                # Add other optional fields if present in CSV and needed by Invoke-RLApi's ShipmentDetails
                CustomerData = if ($headers -contains 'CustomerData') { $row.CustomerData.Trim() } else { $null }
                QuoteType    = if ($headers -contains 'QuoteType') { $row.QuoteType.Trim() } else { 'Domestic' } # Default
                CODAmount    = if ($headers -contains 'CODAmount') { try {[decimal]$row.CODAmount.Trim()} catch {$null} } else { 0.0 }
                ItemWidth    = if ($headers -contains 'ItemWidth') { try {[float]$row.ItemWidth.Trim()} catch {1.0} } else { 1.0 } # Default
                ItemHeight   = if ($headers -contains 'ItemHeight') { try {[float]$row.ItemHeight.Trim()} catch {1.0} } else { 1.0 }
                ItemLength   = if ($headers -contains 'ItemLength') { try {[float]$row.ItemLength.Trim()} catch {1.0} } else { 1.0 }
                DeclaredValue= if ($headers -contains 'DeclaredValue') { try {[decimal]$row.DeclaredValue.Trim()} catch {$null} } else { 0.0 }
            }
            $normalizedData.Add($entry)
        }
        if ($invalidRowCount -gt 0) { Write-Warning " -> Skipped $invalidRowCount R+L rows due to missing/invalid essential data." }
        Write-Host " -> OK: $($normalizedData.Count) R+L rows normalized." -ForegroundColor Green
        return $normalizedData
    } catch { Write-Error "Error processing R+L CSV '$CsvPath': $($_.Exception.Message)"; return $null }
}

function Load-And-Normalize-AverittData {
    param( [Parameter(Mandatory=$true)][string]$CsvPath )
    Write-Host "`nLoading Averitt data from: $(Split-Path $CsvPath -Leaf)" -ForegroundColor Cyan
    try {
        $rawData = Import-Csv -Path $CsvPath -ErrorAction Stop
        Write-Host " -> Rows found: $($rawData.Count)." -ForegroundColor Gray
        if ($rawData.Count -eq 0) { Write-Warning "CSV file is empty."; return @() }

        # Validate presence of essential columns for Averitt API structure
        $requiredCsvColumns = @(
            "ServiceLevel", "PaymentTerms", "PaymentPayer", "PickupDate",
            "OriginCity", "OriginStateProvince", "OriginPostalCode", "OriginCountry",
            "DestinationCity", "DestinationStateProvince", "DestinationPostalCode", "DestinationCountry",
            "Commodity1_Classification", "Commodity1_Weight", "Commodity1_Pieces" # At least one commodity line
        )
        $headers = $rawData[0].PSObject.Properties.Name
        $missingCols = $requiredCsvColumns | Where-Object { $_ -notin $headers }
        if ($missingCols) {
            Write-Error "Averitt CSV file '$CsvPath' is missing required columns for basic API call: $($missingCols -join ', ')"
            return $null
        }
        # Further per-row validation will occur in Invoke-AverittApi when building the payload.
        # This function primarily ensures the CSV can be loaded and has the expected shape.
        Write-Host " -> Successfully loaded $($rawData.Count) rows for Averitt processing (further validation in API call)." -ForegroundColor Green
        return $rawData # Return array of PSCustomObjects, Invoke-AverittApi will parse each
    } catch {
        Write-Error "Error processing Averitt CSV '$CsvPath': $($_.Exception.Message)"; return $null
    }
}

# --- API Call Functions (Carrier Specific) ---

function Invoke-CentralTransportApi {
    [CmdletBinding(DefaultParameterSetName = 'FromShipmentObject')]
    param(
        [Parameter(Mandatory, ParameterSetName='FromShipmentObject')] [PSCustomObject]$ShipmentData,
        [Parameter(Mandatory, ParameterSetName='FromShipmentObject')] [hashtable]$KeyData,
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')] [string]$ApiKey, # Corresponds to accessCode
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')] [string]$OriginZip,
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')] [string]$DestinationZip,
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')] [double]$Weight, # Input can be double for flexibility
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')] [string]$FreightClass,
        [Parameter(Mandatory, ParameterSetName='FromIndividualParams')] [string]$customerNumber,
        [Parameter(ParameterSetName='FromIndividualParams')] [string]$Accessorials = $null # Example, not used by CTX byClass
    )
    $accessCodeToUse = $null; $customerNumberToUse = $null; $originZipToUse = $null; $destZipToUse = $null; $weightToUse = $null; $classToUse = $null
    $tariffNameForLog = "UnknownTariff_CTX" # Default for logging
    $localKeyDataForLog = $null # For consistent tariff name logging

    if ($PSCmdlet.ParameterSetName -eq 'FromShipmentObject') {
        $localKeyDataForLog = $KeyData
        $tariffNameForLog = if ($KeyData.ContainsKey('Name')) { $KeyData.Name } else { $KeyData.TariffFileName | Split-Path -Leaf }
        try {
            $originZipToUse = $ShipmentData.'Origin Postal Code'
            $destZipToUse = $ShipmentData.'Destination Postal Code'
            $weightToUse = [decimal]$ShipmentData.'Total Weight' # API expects numeric
            $classToUse = [string]$ShipmentData.'Freight Class 1'
            if ($KeyData.ContainsKey('accessCode')) { $accessCodeToUse = $KeyData.accessCode } else { throw "'accessCode' missing from KeyData." }
            if ($KeyData.ContainsKey('customerNumber')) { $customerNumberToUse = $KeyData.customerNumber } else { throw "'customerNumber' missing from KeyData." }
            if ([string]::IsNullOrWhiteSpace($accessCodeToUse) -or [string]::IsNullOrWhiteSpace($customerNumberToUse)) { throw "Credentials (accessCode/customerNumber) are empty in KeyData."}
        } catch { Write-Warning "Central API: Parameter extraction failed from Shipment Object for Tariff '$tariffNameForLog'. Error: $($_.Exception.Message)"; return $null }
    } elseif ($PSCmdlet.ParameterSetName -eq 'FromIndividualParams') {
         $accessCodeToUse = $ApiKey; $customerNumberToUse = $customerNumber; $originZipToUse = $OriginZip
         $destZipToUse = $DestinationZip; $weightToUse = [decimal]$Weight; $classToUse = $FreightClass # Convert weight
         $tariffNameForLog = "SingleQuoteCall_CTX"
         # If KeyData is somehow passed (though not part of this param set), use its name for logging
         if ($PSBoundParameters.ContainsKey('KeyData') -and $PSBoundParameters['KeyData'] -is [hashtable] -and $PSBoundParameters['KeyData'].ContainsKey('Name')) {
             $tariffNameForLog = $PSBoundParameters['KeyData'].Name
         }
    } else { Write-Error "Central API: Internal error - Invalid parameter set '$($PSCmdlet.ParameterSetName)'."; return $null }

    # Unified Validation
    $missingParams = @()
    if ([string]::IsNullOrWhiteSpace($originZipToUse) -or -not ($originZipToUse -match '^\d{5}')) { $missingParams += "OriginZip ('$originZipToUse')" }
    if ([string]::IsNullOrWhiteSpace($destZipToUse) -or -not ($destZipToUse -match '^\d{5}')) { $missingParams += "DestinationZip ('$destZipToUse')" }
    if ($null -eq $weightToUse -or $weightToUse -le 0) { $missingParams += "Weight ('$($weightToUse)')" }
    if ([string]::IsNullOrWhiteSpace($classToUse)) { $missingParams += "FreightClass" }
    if ([string]::IsNullOrWhiteSpace($accessCodeToUse)) { $missingParams += "AccessCode (APIKey)" }
    if ([string]::IsNullOrWhiteSpace($customerNumberToUse)) { $missingParams += "CustomerNumber" }
    if ($missingParams.Count -gt 0) {
        Write-Warning "Central API Skip for Tariff '$tariffNameForLog': Missing/invalid required data: $($missingParams -join ', ')."
        return $null
    }

    $rateItemsArray = @( @{ id = 1; weight = $weightToUse; itemClass = $classToUse } ) # API expects numeric weight and string class
    $payload = @{
        accessCode = $accessCodeToUse
        request = @{
            originZipCode = $originZipToUse; destinationZipCode = $destZipToUse
            customerNumber = [string]$customerNumberToUse # Ensure string for JSON
            pickupDate = (Get-Date -Format 'MM/dd/yyyy'); customerRole = "shipper"
            rateItems = $rateItemsArray; useDefaultTariff = $false # Assuming you always provide specific tariff via customerNumber/accessCode
        }
    } | ConvertTo-Json -Depth 5
    $headers = @{ 'Content-Type' = 'application/json' }
    Write-Verbose "Calling Central API for Tariff '$tariffNameForLog': $($script:centralApiUri)"

    try {
        if ([string]::IsNullOrWhiteSpace($script:centralApiUri)) { throw "Central API URI (script:centralApiUri) is not defined or empty."}
        $response = Invoke-RestMethod -Uri $script:centralApiUri -Method Post -Headers $headers -Body $payload -ErrorAction Stop -TimeoutSec 30
        Write-Verbose "Central API OK for Tariff '$tariffNameForLog'."

        $totalChargeValue = $null
        if ($response -ne $null -and $response.PSObject.Properties.Name -contains 'rateTotal') {
            $totalChargeValue = $response.rateTotal
        }

        if ($totalChargeValue -ne $null) {
            try { $cleanedRate = $totalChargeValue -replace '[$,]'; return [decimal]$cleanedRate }
            catch { Write-Warning "Central API Convert Fail for Tariff '$tariffNameForLog': Cannot convert rate '$totalChargeValue' to decimal. Error: $($_.Exception.Message)"; return $null }
        } else { Write-Warning "Central API Response for Tariff '$tariffNameForLog' missing 'rateTotal' or response was null."; return $null }
    } catch {
         $errMsg = $_.Exception.Message; $statusCode = "N/A"; $eBody = "N/A"
         if ($_.Exception.Response) {
             try {$statusCode = $_.Exception.Response.StatusCode.value__} catch{}
             try { $stream = $_.Exception.Response.GetResponseStream(); $reader = New-Object System.IO.StreamReader($stream); $eBody = $reader.ReadToEnd(); $reader.Close(); $stream.Close() }
             catch {$eBody="(Error reading response body: $($_.Exception.Message))"}
         }
         $truncatedBody = if ($eBody.Length -gt 300) { $eBody.Substring(0, 300) + "..." } else { $eBody }
         Write-Warning "Central API FAIL for Tariff '$tariffNameForLog'. Error: $errMsg (HTTP Status: $statusCode) Response: $truncatedBody"; return $null
    }
}

function Invoke-SAIAApi {
    param(
        [Parameter(Mandatory=$true)] [hashtable]$KeyData, # Contains UserID/Password or RQKey, AccountCode, Name
        [Parameter(Mandatory=$true)] [string]$OriginZip,
        [Parameter(Mandatory=$true)] [string]$DestinationZip,
        [Parameter(Mandatory=$true)] [string]$OriginCity,
        [Parameter(Mandatory=$true)] [string]$OriginState, # 2-letter code
        [Parameter(Mandatory=$true)] [string]$DestinationCity,
        [Parameter(Mandatory=$true)] [string]$DestinationState, # 2-letter code
        [Parameter(Mandatory=$true)] [decimal]$Weight,     # Overall weight
        [Parameter(Mandatory=$true)] [string]$Class,        # Primary class
        [Parameter(Mandatory=$false)] [object]$Details = $null # Optional: Can be a PSCustomObject with units, length, width, height for the single commodity
    )
    $saiaUserID = $null; $saiaPassword = $null; $saiaRQKey = $null; $accountCodeToUse = $null
    $tariffNameForLog = if ($KeyData.ContainsKey('Name')) { $KeyData.Name } else { $KeyData.TariffFileName | Split-Path -Leaf }

    try {
        if ($KeyData.ContainsKey('UserID')) { $saiaUserID = $KeyData.UserID }
        if ($KeyData.ContainsKey('Password')) { $saiaPassword = $KeyData.Password }
        if ($KeyData.ContainsKey('RQKey')) { $saiaRQKey = $KeyData.RQKey }
        if ($KeyData.ContainsKey('AccountCode')) { $accountCodeToUse = $KeyData.AccountCode }

        # Construct 'details' array for SAIA API (expects an array, even for one item)
        $detailsArray = @()
        if ($Weight -gt 0 -and -not [string]::IsNullOrWhiteSpace($Class)) {
            $itemWeightInt = 0; $itemClassDouble = 0.0; $itemUnits = 1; $itemLength = 1.0; $itemWidth = 1.0; $itemHeight = 1.0
            try { $itemWeightInt = [int]$Weight } catch { Write-Warning "SAIA API: Cannot convert Weight '$Weight' to integer for details array (Tariff: $tariffNameForLog). Using 0." }
            if (-not [double]::TryParse($Class, [ref]$itemClassDouble)) { Write-Warning "SAIA API: Could not parse Class '$Class' as double for details array (Tariff: $tariffNameForLog). Using 0.0." }

            if ($null -ne $Details -and $Details -is [PSCustomObject]) { # Check if optional $Details object was passed
                 if($Details.PSObject.Properties.Name -contains 'units' -and $null -ne $Details.units) { try { $itemUnits = [int]$Details.units; if($itemUnits -le 0) {$itemUnits = 1} } catch { $itemUnits=1} }
                 if($Details.PSObject.Properties.Name -contains 'length' -and $null -ne $Details.length) { try { $itemLength = [double]$Details.length } catch {$itemLength=1.0} }
                 if($Details.PSObject.Properties.Name -contains 'width' -and $null -ne $Details.width) { try { $itemWidth = [double]$Details.width } catch {$itemWidth=1.0} }
                 if($Details.PSObject.Properties.Name -contains 'height' -and $null -ne $Details.height) { try { $itemHeight = [double]$Details.height } catch {$itemHeight=1.0} }
            }
            if ($itemWeightInt -gt 0) { # Only add if weight is valid
                $detailsArray += @{ length=[double]$itemLength; width=[double]$itemWidth; height=[double]$itemHeight; weight=$itemWeightInt; class=$itemClassDouble; units=$itemUnits }
            }
        }
    } catch { Write-Warning "SAIA API: Parameter extraction/preparation failed for Tariff '$tariffNameForLog'. Error: $($_.Exception.Message)"; return $null }

    # Unified Validation
    $missingParams = @()
    if (-not ($OriginZip -match '^\d{5}')) { $missingParams += "OriginZip ('$OriginZip')" }
    if (-not ($DestinationZip -match '^\d{5}')) { $missingParams += "DestinationZip ('$DestinationZip')" }
    if ([string]::IsNullOrWhiteSpace($OriginCity)) { $missingParams += "OriginCity" }
    if (-not ($OriginState -match '^[A-Za-z]{2}$')) { $missingParams += "OriginState ('$OriginState')" }
    if ([string]::IsNullOrWhiteSpace($DestinationCity)) { $missingParams += "DestinationCity" }
    if (-not ($DestinationState -match '^[A-Za-z]{2}$')) { $missingParams += "DestinationState ('$DestinationState')" }
    if ($Weight -le 0) { $missingParams += "Weight ('$Weight')" }
    if ([string]::IsNullOrWhiteSpace($Class)) { $missingParams += "Class" }
    if ($detailsArray.Count -eq 0) { $missingParams += "Commodity Details (Weight/Class invalid or missing)"}
    if ([string]::IsNullOrWhiteSpace($saiaRQKey) -and ([string]::IsNullOrWhiteSpace($saiaUserID) -or [string]::IsNullOrWhiteSpace($saiaPassword))) { $missingParams += "Credentials (RQKey or UserID/Password pair)" }
    if ($missingParams.Count -gt 0) { Write-Warning "SAIA API Skip for Tariff '$tariffNameForLog': Missing/invalid data: $($missingParams -join ', ')."; return $null }

    $calculatedTotalCube = 0.0 # SAIA API requires totalCube
    try {
        $totalVolumeInches = 0.0
        foreach($item in $detailsArray) {
            if ($item.length -gt 0 -and $item.width -gt 0 -and $item.height -gt 0 -and $item.units -gt 0) {
                 $totalVolumeInches += ($item.length * $item.width * $item.height * $item.units)
            }
        }
        if ($totalVolumeInches -gt 0) { $calculatedTotalCube = [Math]::Round($totalVolumeInches / 1728, 2) } # Cubic inches to cubic feet
    } catch { Write-Warning "SAIA API: Could not calculate totalCube for Tariff '$tariffNameForLog'. Error: $($_.Exception.Message)" }

    $payloadObject = [ordered]@{
        userID = if(-not [string]::IsNullOrWhiteSpace($saiaUserID)) { $saiaUserID } else { $null } # Omit if empty
        password = if(-not [string]::IsNullOrWhiteSpace($saiaPassword)) { $saiaPassword } else { $null } # Omit if empty
        payer = "Shipper"; pickUpDate = (Get-Date -Format 'yyyy-MM-dd')
        origin = @{ city = $OriginCity; state = $OriginState.ToUpper(); zipcode = $OriginZip }
        destination = @{ city = $DestinationCity; state = $DestinationState.ToUpper(); zipcode = $DestinationZip }
        weightUnits = "LBS"; measurementUnit = "IN"; totalCube = $calculatedTotalCube; totalCubeUnits = "CUFT"; details = $detailsArray
    }
    if (-not [string]::IsNullOrWhiteSpace($accountCodeToUse)) { # Conditionally add AccountCode
        $payloadObject.origin.accountCode = $accountCodeToUse
        $payloadObject.destination.accountCode = $accountCodeToUse
    }
    # Remove null credential fields before converting to JSON
    if ($null -eq $payloadObject.userID) { $payloadObject.Remove('userID') }
    if ($null -eq $payloadObject.password) { $payloadObject.Remove('password') }

    $payload = $payloadObject | ConvertTo-Json -Depth 10
    $headers = @{ 'Content-Type' = 'application/json'; 'Cache-Control' = 'no-cache' }
    if (-not [string]::IsNullOrWhiteSpace($saiaRQKey)) { $headers.'RQ-Key' = $saiaRQKey }
    Write-Verbose "Calling SAIA API for Tariff '$tariffNameForLog': $($script:saiaApiUri)"

    try {
        if ([string]::IsNullOrWhiteSpace($script:saiaApiUri)) { throw "SAIA API URI (script:saiaApiUri) is not defined." }
        if (-not $headers.ContainsKey('RQ-Key') -and ($null -eq $payloadObject.userID -or $null -eq $payloadObject.password) ) {
             throw "Cannot call SAIA API for Tariff '$tariffNameForLog': No RQ-Key header provided AND UserID/Password pair is incomplete/missing."
        }
        $response = Invoke-RestMethod -Uri $script:saiaApiUri -Method Post -Headers $headers -Body $payload -ErrorAction Stop -TimeoutSec 30
        Write-Verbose "SAIA API OK for Tariff '$tariffNameForLog'."
        $totalChargeValue = $null
        if ($response -and $response.PSObject.Properties.Name -contains 'rateDetails' -and $response.rateDetails -and $response.rateDetails.PSObject.Properties.Name -contains 'totalInvoice') {
            $totalChargeValue = $response.rateDetails.totalInvoice
        }
        if ($totalChargeValue -ne $null) {
            try { $cleanedRate = $totalChargeValue -replace '[$,]'; return [decimal]$cleanedRate }
            catch { Write-Warning "SAIA API Convert Fail for Tariff '$tariffNameForLog': Cannot convert rate '$totalChargeValue' to decimal. Error: $($_.Exception.Message)"; return $null }
        } else { Write-Warning "SAIA API Response for Tariff '$tariffNameForLog' missing 'rateDetails.totalInvoice' or structure invalid."; return $null }
    } catch {
        $errMsg = $_.Exception.Message; $statusCode = "N/A"; $eBody = "N/A"
        if ($_.Exception.Response) {
             try {$statusCode = $_.Exception.Response.StatusCode.value__} catch{}
             try { $stream = $_.Exception.Response.GetResponseStream(); $reader = New-Object System.IO.StreamReader($stream); $eBody = $reader.ReadToEnd(); $reader.Close(); $stream.Close() }
             catch {$eBody="(Error reading response body: $($_.Exception.Message))"}
        }
        $truncatedBody = if ($eBody.Length -gt 300) { $eBody.Substring(0, 300) + "..." } else { $eBody }
        Write-Warning "SAIA API FAIL for Tariff '$tariffNameForLog'. Error: $errMsg (HTTP Status: $statusCode) Response: $truncatedBody"; return $null
     }
}

function Invoke-RLApi {
    param(
        [Parameter(Mandatory=$true)] [hashtable]$KeyData, # Contains APIKey, Name
        [Parameter(Mandatory=$true)] [string]$OriginZip,
        [Parameter(Mandatory=$true)] [string]$DestinationZip,
        [Parameter(Mandatory=$true)] [decimal]$Weight, # Overall weight
        [Parameter(Mandatory=$true)] [string]$Class,    # Primary class
        [Parameter(Mandatory=$false)] [PSCustomObject]$ShipmentDetails = $null # Optional: For City, State, Dims, DeclaredValue etc.
    )
    $tariffNameForLog = if ($KeyData.ContainsKey('Name')) { $KeyData.Name } else { $KeyData.TariffFileName | Split-Path -Leaf }
    $apiKeyToUse = $null; $customerDataToUse = $null # CustomerData can be in KeyData or ShipmentDetails

    if ($KeyData.ContainsKey('APIKey')) { $apiKeyToUse = $KeyData.APIKey }
    if ($KeyData.ContainsKey('CustomerData')) { $customerDataToUse = $KeyData.CustomerData } # Default from key file

    # Unified Validation
    $missingParams = @()
    if (-not ($OriginZip -match '^\d{5}')) { $missingParams += "OriginZip ('$OriginZip')" }
    if (-not ($DestinationZip -match '^\d{5}')) { $missingParams += "DestinationZip ('$DestinationZip')" }
    if ($Weight -le 0) { $missingParams += "Weight ('$Weight')" }
    if ([string]::IsNullOrWhiteSpace($Class)) { $missingParams += "Class" }
    if ([string]::IsNullOrWhiteSpace($apiKeyToUse)) { $missingParams += "APIKey (from KeyData)" }
    # R+L API requires City/State. Get them from $ShipmentDetails or fail if not present.
    $OriginCityToUse = if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'OriginCity' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.OriginCity)) { $ShipmentDetails.OriginCity } else { $null }
    $OriginStateToUse = if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'OriginState' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.OriginState)) { $ShipmentDetails.OriginState.ToUpper() } else { $null }
    $DestinationCityToUse = if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'DestinationCity' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.DestinationCity)) { $ShipmentDetails.DestinationCity } else { $null }
    $DestinationStateToUse = if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'DestinationState' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.DestinationState)) { $ShipmentDetails.DestinationState.ToUpper() } else { $null }

    if ([string]::IsNullOrWhiteSpace($OriginCityToUse)) { $missingParams += "OriginCity (from ShipmentDetails)"}
    if (-not ($OriginStateToUse -match '^[A-Z]{2}$')) { $missingParams += "OriginState (from ShipmentDetails: '$OriginStateToUse')"}
    if ([string]::IsNullOrWhiteSpace($DestinationCityToUse)) { $missingParams += "DestinationCity (from ShipmentDetails)"}
    if (-not ($DestinationStateToUse -match '^[A-Z]{2}$')) { $missingParams += "DestinationState (from ShipmentDetails: '$DestinationStateToUse')"}
    if ($missingParams.Count -gt 0) { Write-Warning "R+L API Skip for Tariff '$tariffNameForLog': Missing/invalid data: $($missingParams -join ', ')." ; return $null }

    # Optional fields from ShipmentDetails or defaults
    $OriginCountryCodeToUse = if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'OriginCountryCode' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.OriginCountryCode)) { $ShipmentDetails.OriginCountryCode } else { 'USA' }
    $DestinationCountryCodeToUse = if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'DestinationCountryCode' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.DestinationCountryCode)) { $ShipmentDetails.DestinationCountryCode } else { 'USA' }
    $QuoteTypeToUse = if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'QuoteType' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.QuoteType)) { $ShipmentDetails.QuoteType } else { 'Domestic' }
    $CODAmountToUse = if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'CODAmount' -and $null -ne $ShipmentDetails.CODAmount) { try { [decimal]$ShipmentDetails.CODAmount } catch {0.0} } else { 0.0 }
    $DeclaredValueToUse = if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'DeclaredValue' -and $null -ne $ShipmentDetails.DeclaredValue) { try { [decimal]$ShipmentDetails.DeclaredValue } catch {0.0} } else { 0.0 }
    $ItemWidthToUse = if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'ItemWidth' -and $null -ne $ShipmentDetails.ItemWidth) { try { [float]$ShipmentDetails.ItemWidth } catch {1.0} } else { 1.0 }
    $ItemHeightToUse = if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'ItemHeight' -and $null -ne $ShipmentDetails.ItemHeight) { try { [float]$ShipmentDetails.ItemHeight } catch {1.0} } else { 1.0 }
    $ItemLengthToUse = if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'ItemLength' -and $null -ne $ShipmentDetails.ItemLength) { try { [float]$ShipmentDetails.ItemLength } catch {1.0} } else { 1.0 }
    if ($null -ne $ShipmentDetails -and $ShipmentDetails.PSObject.Properties.Name -contains 'CustomerData' -and -not [string]::IsNullOrWhiteSpace($ShipmentDetails.CustomerData)) { $customerDataToUse = $ShipmentDetails.CustomerData } # Override from ShipmentDetails if present

    $ItemWeightToUse = try {[float]$Weight} catch { Write-Warning "R+L API: Could not convert Weight '$Weight' to float for item payload (Tariff: $tariffNameForLog)."; 0.0 }
    if ($ItemWeightToUse -le 0) { Write-Warning "R+L API Skip for Tariff '$tariffNameForLog': Invalid Item Weight ($ItemWeightToUse) after conversion."; return $null }

    $soapEndpoint = $script:rlApiUri; if ([string]::IsNullOrWhiteSpace($soapEndpoint)) { throw "R+L API URI (script:rlApiUri) is not defined." }
    $soapAction = "http://www.rlcarriers.com/GetRateQuote"; $tnsNamespace = "http://www.rlcarriers.com/"; $soapNamespace = "http://schemas.xmlsoap.org/soap/envelope/"
    function Escape-Xml ($stringToEscape) { if ($null -eq $stringToEscape) { return '' }; return [System.Security.SecurityElement]::Escape($stringToEscape) } # Renamed param
    $soapRequestBody = @"
<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="$soapNamespace"><soap:Body><tns:GetRateQuote xmlns:tns="$tnsNamespace">
<tns:APIKey>$(Escape-Xml $apiKeyToUse)</tns:APIKey><tns:request>
<tns:CustomerData>$(Escape-Xml $customerDataToUse)</tns:CustomerData><tns:QuoteType>$(Escape-Xml $QuoteTypeToUse)</tns:QuoteType><tns:CODAmount>$CODAmountToUse</tns:CODAmount>
<tns:Origin><tns:City>$(Escape-Xml $OriginCityToUse)</tns:City><tns:StateOrProvince>$(Escape-Xml $OriginStateToUse)</tns:StateOrProvince><tns:ZipOrPostalCode>$(Escape-Xml $OriginZip)</tns:ZipOrPostalCode><tns:CountryCode>$(Escape-Xml $OriginCountryCodeToUse)</tns:CountryCode></tns:Origin>
<tns:Destination><tns:City>$(Escape-Xml $DestinationCityToUse)</tns:City><tns:StateOrProvince>$(Escape-Xml $DestinationStateToUse)</tns:StateOrProvince><tns:ZipOrPostalCode>$(Escape-Xml $DestinationZip)</tns:ZipOrPostalCode><tns:CountryCode>$(Escape-Xml $DestinationCountryCodeToUse)</tns:CountryCode></tns:Destination>
<tns:Items><tns:Item><tns:Class>$(Escape-Xml $Class)</tns:Class><tns:Weight>$ItemWeightToUse</tns:Weight><tns:Width>$ItemWidthToUse</tns:Width><tns:Height>$ItemHeightToUse</tns:Height><tns:Length>$ItemLengthToUse</tns:Length></tns:Item></tns:Items>
<tns:DeclaredValue>$DeclaredValueToUse</tns:DeclaredValue></tns:request></tns:GetRateQuote></soap:Body></soap:Envelope>
"@
    $headers = @{ "Content-Type" = "text/xml; charset=utf-8"; "SOAPAction" = "`"$soapAction`"" }
    Write-Verbose "Calling R+L API for Tariff '$tariffNameForLog': $soapEndpoint"
    try {
        $ProgressPreference = 'SilentlyContinue' # Suppress Invoke-WebRequest progress bar
        $response = Invoke-WebRequest -Uri $soapEndpoint -Method Post -Headers $headers -Body $soapRequestBody -UseBasicParsing -ErrorAction Stop -TimeoutSec 30
        [xml]$responseXml = $response.Content; $nsManager = New-Object System.Xml.XmlNamespaceManager($responseXml.NameTable)
        $nsManager.AddNamespace("soap", $soapNamespace); $nsManager.AddNamespace("rl", $tnsNamespace)
        $faultNode = $responseXml.SelectSingleNode("/soap:Envelope/soap:Body/soap:Fault", $nsManager)
        if ($faultNode) { $faultString = $faultNode.SelectSingleNode("faultstring", $nsManager).InnerText; $faultCode = $faultNode.SelectSingleNode("faultcode", $nsManager).InnerText; Write-Warning "R+L API Fault for Tariff '$($tariffNameForLog)': Code='$faultCode', String='$faultString'"; return $null }
        $rateQuoteResult = $responseXml.SelectSingleNode("/soap:Envelope/soap:Body/rl:GetRateQuoteResponse/rl:GetRateQuoteResult", $nsManager)
        if ($rateQuoteResult) {
            $quoteDetails = $rateQuoteResult.SelectSingleNode("rl:Result", $nsManager)
            if ($quoteDetails) {
                $netChargeEntry = $quoteDetails.SelectSingleNode("rl:Charges/rl:Charge[rl:Type='NET']/rl:Amount", $nsManager)
                if ($netChargeEntry) {
                    $totalChargeValue = $netChargeEntry.InnerText; try { $cleanedRate = $totalChargeValue -replace '[$,]'; return [decimal]$cleanedRate }
                    catch { Write-Warning "R+L API Convert Fail for Tariff '$($tariffNameForLog)': Cannot convert rate '$totalChargeValue' to decimal. Error: $($_.Exception.Message)"; return $null }
                } else {
                    $errorMsgNode = $quoteDetails.SelectSingleNode("rl:Errors/rl:string", $nsManager)
                    if($errorMsgNode){ Write-Warning "R+L API Error in Response for Tariff '$($tariffNameForLog)': $($errorMsgNode.InnerText)" } else { Write-Warning "R+L API Response for Tariff '$tariffNameForLog' Missing 'NET' charge. Check response XML." }
                    return $null
                }
            } else { Write-Warning "R+L API Response for Tariff '$tariffNameForLog' structure unexpected (Cannot find 'Result' element)."; return $null }
        } else { Write-Warning "R+L API Response for Tariff '$tariffNameForLog' structure unexpected (Cannot find 'GetRateQuoteResult' or 'Fault' element)."; return $null }
    } catch {
        $errMsg = $_.Exception.Message; $statusCode = "N/A"; $eBody = "N/A"
        if ($_.Exception.Response) { try {$statusCode = $_.Exception.Response.StatusCode.value__} catch{}; try { $stream = $_.Exception.Response.GetResponseStream(); $reader = New-Object System.IO.StreamReader($stream); $eBody = $reader.ReadToEnd(); $reader.Close(); $stream.Close() } catch {$eBody="(Error reading response body)"} }
        $truncatedBody = if ($eBody.Length -gt 300) { $eBody.Substring(0, 300) + "..." } else { $eBody }
        Write-Warning "R+L API FAIL for Tariff '$tariffNameForLog'. Error: $errMsg (HTTP Status: $statusCode) Response: $truncatedBody"; return $null
    }
}

function Invoke-AverittApi {
    param(
        [Parameter(Mandatory=$true)][hashtable]$KeyData,         # Contains APIKey, MarginPercent, Name
        [Parameter(Mandatory=$true)][PSCustomObject]$ShipmentData # A single row object from Load-And-Normalize-AverittData
    )

    $apiKeyToUse = $KeyData.APIKey
    $tariffNameForLog = if ($KeyData.ContainsKey('Name')) { $KeyData.Name } else { "Averitt_UnknownTariff" }

    if ([string]::IsNullOrWhiteSpace($apiKeyToUse)) {
        Write-Warning "Averitt API Key missing for tariff '$tariffNameForLog'. Skipping API call."
        return $null
    }

    $RateApiUrl = $script:averittApiUri # From TMS_Config.ps1
    if ([string]::IsNullOrWhiteSpace($RateApiUrl)) { throw "Averitt API URI (script:averittApiUri) is not defined."}

    # Construct RequestBody from $ShipmentData
    $RequestBody = @{
        "service"     = @{ "level" = $ShipmentData.ServiceLevel }
        "payment"     = @{ "terms" = $ShipmentData.PaymentTerms; "payer" = $ShipmentData.PaymentPayer }
        "transit"     = @{ "pickupDate" = $ShipmentData.PickupDate } # Ensure YYYYMMDD or API required format
        "commodities" = @()
        "accessorials"= @{ "codes" = @(); "hazardousContact" = $null; "cod" = $null; "insuranceDetails" = $null; "sortAndSegregateDetails" = $null; "markDetails" = $null }
        "origin"      = @{ "account" = $ShipmentData.OriginAccount; "city" = $ShipmentData.OriginCity; "stateProvince" = $ShipmentData.OriginStateProvince; "postalCode" = $ShipmentData.OriginPostalCode; "country" = $ShipmentData.OriginCountry }
        "destination" = @{ "account" = $ShipmentData.DestinationAccount; "city" = $ShipmentData.DestinationCity; "stateProvince" = $ShipmentData.DestinationStateProvince; "postalCode" = $ShipmentData.DestinationPostalCode; "country" = $ShipmentData.DestinationCountry }
        "billTo"      = @{ "account" = $ShipmentData.BillToAccount; "name" = $ShipmentData.BillToName; "address" = $ShipmentData.BillToAddress; "city" = $ShipmentData.BillToCity; "stateProvince" = $ShipmentData.BillToStateProvince; "postalCode" = $ShipmentData.BillToPostalCode; "country" = $ShipmentData.BillToCountry }
    }

    for ($commIdx = 1; $commIdx -le 5; $commIdx++) {
        $classKey = "Commodity${commIdx}_Classification"; $weightKey = "Commodity${commIdx}_Weight"; $piecesKey = "Commodity${commIdx}_Pieces"
        $lengthKey = "Commodity${commIdx}_Length"; $widthKey = "Commodity${commIdx}_Width"; $heightKey = "Commodity${commIdx}_Height"
        $pkgTypeKey = "Commodity${commIdx}_PackagingType"; $descKey = "Commodity${commIdx}_Description"; $stackableKey = "Commodity${commIdx}_Stackable"

        if ($ShipmentData.PSObject.Properties.Match($classKey) -and (-not [string]::IsNullOrWhiteSpace($ShipmentData.$classKey)) -and
            $ShipmentData.PSObject.Properties.Match($weightKey) -and (-not [string]::IsNullOrWhiteSpace($ShipmentData.$weightKey)) -and
            $ShipmentData.PSObject.Properties.Match($piecesKey) -and (-not [string]::IsNullOrWhiteSpace($ShipmentData.$piecesKey))) {
            
            $commodityWeight = 0.0
            try { $commodityWeight = [decimal]$ShipmentData.$weightKey } catch { Write-Warning "Averitt API: Invalid weight for Commodity ${commIdx} ('$($ShipmentData.$weightKey)') for Tariff '${tariffNameForLog}'. Skipping commodity line."; continue }
            if ($commodityWeight -le 0) { Write-Warning "Averitt API: Non-positive weight for Commodity ${commIdx} ($($commodityWeight)) for Tariff '${tariffNameForLog}'. Skipping commodity line."; continue }

            $commodity = @{
                "classification" = $ShipmentData.$classKey
                "weight"         = $commodityWeight # API likely expects numeric
                "pieces"         = $ShipmentData.$piecesKey
            }
            # Add optional commodity fields
            if ($ShipmentData.PSObject.Properties.Match($lengthKey) -and -not [string]::IsNullOrWhiteSpace($ShipmentData.$lengthKey)) { $commodity.Add("length", ([double]$ShipmentData.$lengthKey)) } # Ensure numeric
            if ($ShipmentData.PSObject.Properties.Match($widthKey) -and -not [string]::IsNullOrWhiteSpace($ShipmentData.$widthKey)) { $commodity.Add("width", ([double]$ShipmentData.$widthKey)) }   # Ensure numeric
            if ($ShipmentData.PSObject.Properties.Match($heightKey) -and -not [string]::IsNullOrWhiteSpace($ShipmentData.$heightKey)) { $commodity.Add("height", ([double]$ShipmentData.$heightKey)) } # Ensure numeric
            if ($ShipmentData.PSObject.Properties.Match($pkgTypeKey) -and -not [string]::IsNullOrWhiteSpace($ShipmentData.$pkgTypeKey)) { $commodity.Add("packagingType", $ShipmentData.$pkgTypeKey) }
            if ($ShipmentData.PSObject.Properties.Match($descKey) -and -not [string]::IsNullOrWhiteSpace($ShipmentData.$descKey)) { $commodity.Add("description", $ShipmentData.$descKey) }
            if ($ShipmentData.PSObject.Properties.Match($stackableKey) -and -not [string]::IsNullOrWhiteSpace($ShipmentData.$stackableKey)) {
                $isStackable = $false
                if ($ShipmentData.$stackableKey -is [boolean]) { $isStackable = $ShipmentData.$stackableKey } # If already boolean
                elseif ($ShipmentData.$stackableKey -eq 'Y' -or $ShipmentData.$stackableKey -eq 'true') { $isStackable = $true }
                $commodity.Add("stackable", $isStackable) 
            }
            $RequestBody.commodities += $commodity
        }
    }
    
    if ($RequestBody.commodities.Count -eq 0) {
        Write-Warning "Averitt API Invoke for Tariff '$tariffNameForLog': No valid commodity lines constructed for shipment (Origin: $($ShipmentData.OriginPostalCode)). Skipping API call."
        return $null
    }

    if ($ShipmentData.PSObject.Properties.Match("AccessorialCodes") -and -not [string]::IsNullOrWhiteSpace($ShipmentData.AccessorialCodes)) {
        $RequestBody.accessorials.codes = @($ShipmentData.AccessorialCodes.Split(',') | ForEach-Object {$_.Trim()} | Where-Object {-not [string]::IsNullOrWhiteSpace($_)})
    }
    if ($ShipmentData.PSObject.Properties.Match("HazardousContactName") -and (-not [string]::IsNullOrWhiteSpace($ShipmentData.HazardousContactName) -or -not [string]::IsNullOrWhiteSpace($ShipmentData.HazardousContactPhone))) {
        $RequestBody.accessorials.hazardousContact = @{ "name" = $ShipmentData.HazardousContactName; "phone" = $ShipmentData.HazardousContactPhone }
    }

    $JsonBody = $RequestBody | ConvertTo-Json -Depth 10 -Compress # Compress for potentially smaller payload
    $FullApiUrl = "$($RateApiUrl)?api_key=$($apiKeyToUse)"
    $Headers = @{ "Content-Type" = "application/json"; "Accept" = "application/json" }
    Write-Verbose "Calling Averitt API for Tariff '$tariffNameForLog': $FullApiUrl"
    # For deep debugging: Write-Verbose "Averitt Request Body for Tariff '$tariffNameForLog': $JsonBody"

    try {
        $apiResponse = Invoke-RestMethod -Uri $FullApiUrl -Method Post -Headers $Headers -Body $JsonBody -ErrorAction Stop -TimeoutSec 60
        if ($apiResponse -and $apiResponse.PSObject.Properties.Name -contains 'totalCharge') {
            Write-Verbose "Averitt API OK for Tariff '$tariffNameForLog'. Total Charge: $($apiResponse.totalCharge)"
            try {
                $chargeAsString = [string]$apiResponse.totalCharge # Ensure it's string before replace
                return [decimal]($chargeAsString -replace '[$,]') # Remove currency symbols
            } catch {
                Write-Warning "Averitt API Convert Fail for Tariff '${tariffNameForLog}': Cannot convert totalCharge '$($apiResponse.totalCharge)' to decimal. Error: $($_.Exception.Message)"
                return $null
            }
        } elseif ($apiResponse -and $apiResponse.PSObject.Properties.Name -contains 'errors' -and $apiResponse.errors) {
            $errorMessages = ($apiResponse.errors | ForEach-Object { "$($_.code): $($_.message)" }) -join "; " # Include error codes
            Write-Warning "Averitt API Error for Tariff '${tariffNameForLog}': ${errorMessages}"
            Write-Verbose "Averitt Full Error Response for Tariff '$tariffNameForLog': $($apiResponse | ConvertTo-Json -Depth 5 -Compress)"
            return $null
        } else {
            Write-Warning "Averitt API Response for Tariff '$tariffNameForLog' structure unexpected. No 'totalCharge' or 'errors' field found."
            Write-Verbose "Averitt Full Unexpected Response for Tariff '$tariffNameForLog': $($apiResponse | ConvertTo-Json -Depth 5 -Compress)"
            return $null
        }
    } catch {
        $errMsg = $_.Exception.Message; $statusCode = "N/A"; $eBody = "N/A"
        if ($_.Exception.Response) {
            try { $statusCode = $_.Exception.Response.StatusCode.value__ } catch { }
            try { $stream = $_.Exception.Response.GetResponseStream(); $reader = New-Object System.IO.StreamReader($stream); $eBody = $reader.ReadToEnd(); $reader.Close(); $stream.Close() }
            catch { $eBody = "(Error reading response body: $($_.Exception.Message))" }
        }
        $truncatedBody = if ($eBody.Length -gt 300) { $eBody.Substring(0, 300) + "..." } else { $eBody }
        Write-Warning "Averitt API FAIL for Tariff '${tariffNameForLog}'. HTTP Error: $errMsg (Status: $statusCode) Response: $truncatedBody"
        return $null
    }
}

# --- Single Quote Calculation Helper Functions ---

function Get-MinimumRate {
    # Finds the minimum cost from a hashtable where keys are TariffNames and values are costs.
    param( [Parameter(Mandatory=$true)][hashtable]$RateResults )
    $lowestCost = $null; $bestTariff = $null
    foreach ($tariffName in $RateResults.Keys) {
         $cost = $RateResults[$tariffName]
         if ($cost -ne $null -and $cost -is [decimal] -and $cost -gt 0) { # Ensure cost is valid decimal
              if ($lowestCost -eq $null -or $cost -lt $lowestCost) { $lowestCost = $cost; $bestTariff = $tariffName }
         }
    }
    if ($lowestCost -ne $null) { return [PSCustomObject]@{ TariffName = $bestTariff; Cost = $lowestCost } }
    else { return $null }
}

function Get-HistoricalAveragePrice {
    # Retrieves historical average price based on criteria.
    param(
        [Parameter(Mandatory=$true)] [string]$OriginZip,
        [Parameter(Mandatory=$true)] [string]$DestinationZip,
        [Parameter(Mandatory=$true)] [double]$Weight, # Can be double for input flexibility
        [Parameter(Mandatory=$true)] [string]$FreightClass
    )
    $histFileName = $Global:HistoricalDataSourceFileName
    if ([string]::IsNullOrWhiteSpace($histFileName)) { Write-Warning "Global:HistoricalDataSourceFileName not set in config. Cannot get historical average."; return $null }
    $histPath = Join-Path $script:shipmentDataFolderPath $histFileName # $script:shipmentDataFolderPath set by main script
    Write-Verbose "Looking up historical average: $OriginZip -> $DestinationZip, Wt:$Weight, Cls:$FreightClass. File: $(Split-Path $histPath -Leaf)"
    $cutoffDate = (Get-Date).AddMonths(-12); $averagePrice = $null
    if (-not (Test-Path $histPath -PathType Leaf)) { Write-Warning "Historical data file missing: $histPath"; return $null }
    if ([string]::IsNullOrWhiteSpace($OriginZip) -or $OriginZip.Length -lt 3 -or [string]::IsNullOrWhiteSpace($DestinationZip) -or $DestinationZip.Length -lt 3) {
        Write-Warning "Historical Lookup Skip: Origin/Destination ZIP too short or missing."; return $null
    }

    $originZip3 = $OriginZip.Substring(0,3); $destinationZip3 = $DestinationZip.Substring(0,3)
    $weightTolerance = $Global:HistoricalWeightTolerance
    if ($null -eq $weightTolerance) { Write-Warning "Global:HistoricalWeightTolerance not set. Defaulting to 0.10 (10%)."; $weightTolerance = 0.10 }
    $minWeight = $Weight * (1.0 - $weightTolerance); $maxWeight = $Weight * (1.0 + $weightTolerance)
    Write-Verbose " -> Historical Weight Range for matching: $minWeight - $maxWeight"

    try {
        $historicalData = Import-Csv -Path $histPath -ErrorAction Stop
        # Define column names expected in the history CSV - these should match your actual file
        $oZipCol='Origin Postal Code'; $dZipCol='Destination Postal Code'; $wtCol='Total Weight'; $clsCol='Freight Class 1'; $prcCol='Price'; $dtCol='Booked Date'
        $requiredHistCols = @($oZipCol, $dZipCol, $wtCol, $clsCol, $prcCol, $dtCol)
        $headerRow = $historicalData[0].PSObject.Properties.Name
        $missingHistCols = $requiredHistCols | Where-Object { $_ -notin $headerRow }
        if ($missingHistCols.Count -gt 0) { Write-Warning "Historical data file '$histPath' is missing required columns: $($missingHistCols -join ', '). Cannot calculate average."; return $null }

        $similarShipments = $historicalData | Where-Object {
            $rowIsValid = $true
            $rowOriginZip = $null; $rowDestZip = $null; $rowWeight = $null; $rowClass = $null; $rowPrice = $null; $rowDate = $null
            try { $rowOriginZip = $_.$oZipCol; if([string]::IsNullOrWhiteSpace($rowOriginZip) -or $rowOriginZip.Length -lt 3) {$rowIsValid=$false} } catch {$rowIsValid=$false}
            if($rowIsValid){ try { $rowDestZip = $_.$dZipCol; if([string]::IsNullOrWhiteSpace($rowDestZip) -or $rowDestZip.Length -lt 3) {$rowIsValid=$false} } catch {$rowIsValid=$false} }
            if($rowIsValid){ try { $rowClass = $_.$clsCol; if([string]::IsNullOrWhiteSpace($rowClass)) {$rowIsValid=$false} } catch {$rowIsValid=$false} }
            if($rowIsValid){ try { $rowWeight = [double]$_.$wtCol; if($rowWeight -le 0) {$rowIsValid=$false} } catch {$rowIsValid=$false} }
            if($rowIsValid){ try { $rowPrice = [double]$_.$prcCol; if($rowPrice -le 0) {$rowIsValid=$false} } catch {$rowIsValid=$false} }
            if($rowIsValid){ try { $rowDate = [datetime]$_.$dtCol } catch {$rowIsValid=$false} }

            if ($rowIsValid) {
                ($rowDate -ge $cutoffDate) -and
                ($rowOriginZip.Substring(0, 3) -eq $originZip3) -and
                ($rowDestZip.Substring(0, 3) -eq $destinationZip3) -and
                ($rowClass -eq $FreightClass) -and # Assumes exact class match
                ($rowWeight -ge $minWeight) -and ($rowWeight -le $maxWeight)
            } else { $false } # Explicitly return false if row data is invalid
        }

        if ($similarShipments -and $similarShipments.Count -gt 0) {
            $prices = $similarShipments | Select-Object -ExpandProperty $prcCol | ForEach-Object { try { [double]$_ } catch {} } | Where-Object {$_ -ne $null}
            if ($prices.Count -gt 0) {
                $measureResult = $prices | Measure-Object -Average -ErrorAction SilentlyContinue
                if ($measureResult -and $measureResult.PSObject.Properties.Name -contains 'Average') {
                    $averagePrice = [math]::Round($measureResult.Average, 2)
                    Write-Verbose "Found $($prices.Count) historical matches. Average Price: $($averagePrice.ToString("C"))"
                } else { Write-Warning "Could not calculate average from historical matches (Measure-Object failed or 'Average' property missing)." }
            } else { Write-Verbose "No valid prices found in matching historical shipments."}
        } else { Write-Verbose "No similar historical shipments found matching all criteria." }
    } catch { Write-Error "Error processing historical data file '$histPath': $($_.Exception.Message)" }
    return $averagePrice
}

function Calculate-QuotePrice {
    # Calculates the final quote price based on lowest carrier cost, margin, and historical average.
    param(
        [Parameter(Mandatory=$true)] [decimal]$LowestCarrierCost,
        [Parameter(Mandatory=$true)] [string]$OriginZip,
        [Parameter(Mandatory=$true)] [string]$DestinationZip,
        [Parameter(Mandatory=$true)] [double]$Weight,
        [Parameter(Mandatory=$true)] [string]$FreightClass,
        [Parameter(Mandatory=$true)] [double]$MarginPercent # Margin as a percentage, e.g., 15.0 for 15%
    )
    $standardMarginPrice = $null
    $marginDecimal = [decimal]$MarginPercent / 100.0

    if ($LowestCarrierCost -gt 0) {
        if ((1.0 - $marginDecimal) -ne 0) { # Avoid division by zero for 100% margin
             try { $standardMarginPrice = [math]::Round(($LowestCarrierCost / (1.0 - $marginDecimal)), 2) }
             catch { Write-Warning "Error calculating standard margin price: $($_.Exception.Message)" }
        } else { Write-Warning "Cannot calculate standard margin price with a 100% margin." }
    } else { Write-Verbose "Lowest carrier cost is zero or less; standard margin price cannot be calculated based on it." }

    $historicalAveragePrice = Get-HistoricalAveragePrice -OriginZip $OriginZip -DestinationZip $DestinationZip -Weight $Weight -FreightClass $FreightClass
    $finalPriceToQuote = $null; $quoteReason = "Error determining final price."

    if ($standardMarginPrice -ne $null -and $historicalAveragePrice -ne $null -and $historicalAveragePrice -gt 0) {
        if ($historicalAveragePrice -gt $standardMarginPrice) {
             $finalPriceToQuote = $historicalAveragePrice
             $quoteReason = "Historical Average Used (Higher than Standard Margin Price)"
        } else {
             $finalPriceToQuote = $standardMarginPrice
             $quoteReason = "Standard Margin ($($MarginPercent)%) Used (Higher or Equal to Historical Average)"
        }
    } elseif ($standardMarginPrice -ne $null) {
        $finalPriceToQuote = $standardMarginPrice
        $quoteReason = "Standard Margin ($($MarginPercent)%) Used (No Valid Historical Average)"
    } elseif ($historicalAveragePrice -ne $null -and $historicalAveragePrice -gt 0) {
        $finalPriceToQuote = $historicalAveragePrice
        $quoteReason = "Historical Average Used (Standard Margin Calculation Error or Zero Cost)"
    } else {
        Write-Warning "Cannot determine a valid final price. Both Standard Margin Price and Historical Average are invalid or zero."
    }

    $finalPriceRounded = if($finalPriceToQuote -ne $null){[math]::Round($finalPriceToQuote, 2)}else{$null}

    return [PSCustomObject]@{
        LowestCarrierCost    = $LowestCarrierCost
        StandardMarginPrice  = $standardMarginPrice
        HistoricalAvgPrice   = $historicalAveragePrice
        MarginUsedPercent    = $MarginPercent
        FinalPrice           = $finalPriceRounded
        Reason               = $quoteReason
    }
}

function Write-QuoteToHistory {
    # Logs a generated quote to a CSV file.
     param(
        [Parameter(Mandatory=$true)] [string]$Carrier,
        [Parameter(Mandatory=$true)] [string]$TariffName,
        [Parameter(Mandatory=$true)] [string]$OriginZip,
        [Parameter(Mandatory=$true)] [string]$DestinationZip,
        [Parameter(Mandatory=$true)] [double]$Weight,
        [Parameter(Mandatory=$true)] [string]$FreightClass,
        [Parameter(Mandatory=$true)] [decimal]$LowestCost,       # Carrier's cost
        [Parameter(Mandatory=$true)] [decimal]$FinalQuotedPrice, # Price quoted to customer
        [Parameter(Mandatory=$true)] [string]$QuoteTimestamp     # e.g., Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
     )
    $logFileName = $Global:HistoricalQuotesLogFileName
    if ([string]::IsNullOrWhiteSpace($logFileName)) { Write-Warning "Global:HistoricalQuotesLogFileName not set. Cannot log quote."; return }
    $logPath = Join-Path $script:shipmentDataFolderPath $logFileName # $script:shipmentDataFolderPath set by main script

    # Basic validation before logging
    if ([string]::IsNullOrWhiteSpace($OriginZip) -or $OriginZip.Length -lt 3 -or [string]::IsNullOrWhiteSpace($DestinationZip) -or $DestinationZip.Length -lt 3){ Write-Warning "Quote Log Skip: Origin/Destination ZIP too short or missing."; return }
    if ($Weight -le 0) { Write-Warning "Quote Log Skip: Invalid Weight ($Weight)."; return }
    if ($LowestCost -le 0) { Write-Warning "Quote Log Skip: Invalid Lowest Carrier Cost ($LowestCost)."; return } # It's possible to quote above a $0 cost, but log if cost itself is invalid
    if ($FinalQuotedPrice -le 0) { Write-Warning "Quote Log Skip: Invalid Final Quoted Price ($FinalQuotedPrice)."; return }

    $originZip3 = $OriginZip.Substring(0,3); $destinationZip3 = $DestinationZip.Substring(0,3)
    $logEntry=[PSCustomObject]@{
        Timestamp            = $QuoteTimestamp
        Carrier              = $Carrier
        Tariff               = $TariffName
        OriginZip3           = $originZip3
        DestZip3             = $destinationZip3
        Weight               = $Weight
        FreightClass         = $FreightClass
        LowestCost           = $LowestCost       # Store as decimal
        FinalQuotedPrice     = $FinalQuotedPrice # Store as decimal
        OriginZipFull        = $OriginZip
        DestinationZipFull   = $DestinationZip
    }

    try {
        Ensure-DirectoryExists -Path (Split-Path $logPath -Parent) # Ensure reports/shipmentData folder exists
        $fileExists = Test-Path $logPath -PathType Leaf
        if ($fileExists) {
             $logEntry | Export-Csv -Path $logPath -NoTypeInformation -Append -Encoding UTF8 -ErrorAction Stop
        } else { # File does not exist, write header
             $logEntry | Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
        }
        Write-Verbose "Quote successfully logged to: $logPath"
    } catch {
        Write-Error "Failed to log quote to '$logPath': $($_.Exception.Message)"
    }
}

# --- Report Path Generation ---
function Get-ReportPath {
    # Centralizes report filename generation and folder creation.
    param(
        [Parameter(Mandatory=$true)][string]$BaseDir,          # e.g., $script:reportsBaseFolderPath
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$Carrier,          # e.g., "Central", "SAIA", "Averitt"
        [Parameter(Mandatory=$true)][string]$ReportType,       # e.g., "Comparison", "AvgMarginCalc", "MarginForASP"
        [string]$FilePrefix = $null,                           # Optional prefix like Key1_vs_Key2
        [string]$FileExtension = "txt"
    )

    $userReportsFolder = Join-Path -Path $BaseDir -ChildPath $Username
    try {
         Ensure-DirectoryExists -Path $userReportsFolder # Ensure user's specific report folder exists
    } catch {
         Write-Error "Failed to ensure user report directory '$userReportsFolder' exists. Cannot generate report path."
         return $null
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $safeCarrierName = $Carrier -replace '[^a-zA-Z0-9_-]', '' # Sanitize for filename
    $safeReportTypeName = $ReportType -replace '[^a-zA-Z0-9_-]', '' # Sanitize
    $safeFilePrefix = if ($FilePrefix) { ($FilePrefix -replace '[^a-zA-Z0-9_-]', '').TrimStart('_').TrimEnd('_') + "_" } else { "" }

    $reportFileName = "{0}_{1}_{2}{3}.{4}" -f $safeCarrierName, $safeReportTypeName, $safeFilePrefix, $timestamp, $FileExtension
    $fullReportPath = Join-Path -Path $userReportsFolder -ChildPath $reportFileName

    Write-Verbose "Generated report path: $fullReportPath"
    return $fullReportPath
}

Write-Verbose "TMS Helper Functions loaded."
