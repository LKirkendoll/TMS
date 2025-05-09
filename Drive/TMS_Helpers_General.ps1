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

# --- Other General Helper Functions (Ensure-DirectoryExists, Get-PermittedKeys, etc.) would be here ---
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
                if($keyEntry.Value -is [hashtable] -and $keyEntry.Value.TariffFileName -eq $allowedName){
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
        [string]$InitialDirectory = $PSScriptRoot
    )
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = $DialogTitle
    $openFileDialog.InitialDirectory = $InitialDirectory
    $openFileDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    }
    return $null
}

Function Open-FileExplorer {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    if (Test-Path $Path) {
        Invoke-Item $Path
    } else {
        Write-Warning "Path not found: $Path"
    }
}

Write-Verbose "TMS General Helper Functions loaded."
