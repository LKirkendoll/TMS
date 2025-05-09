# TMS_Auth.ps1
# Description: Handles user authentication and loading of customer profiles.
#              This file should be dot-sourced by the main script.
# Usage: . .\TMS_Auth.ps1

# --- Hashing Function (Placeholder - REPLACE with a strong, salted hashing algorithm like Argon2 or scrypt) ---
# For demonstration, this uses a simple SHA256 hash. THIS IS NOT SECURE FOR PRODUCTION.
function Get-PasswordHash {
    param([Parameter(Mandatory=$true)][string]$Password)
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($Password)
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $hashedBytes = $sha256.ComputeHash($bytes)
        $sha256.Dispose() # Dispose of the cryptographic object
        return [System.BitConverter]::ToString($hashedBytes).Replace("-", "").ToLowerInvariant()
    } catch {
        Write-Error "Error generating password hash: $($_.Exception.Message)"
        return $null
    }
}

function Test-PasswordHash {
    param(
        [Parameter(Mandatory=$true)][string]$PasswordPlainText,
        [Parameter(Mandatory=$true)][string]$StoredHash
    )
    if ([string]::IsNullOrWhiteSpace($PasswordPlainText) -or [string]::IsNullOrWhiteSpace($StoredHash)) {
        Write-Warning "Password or stored hash is empty."
        return $false
    }
    $computedHash = Get-PasswordHash -Password $PasswordPlainText
    if ($null -eq $computedHash) { return $false } # Error in Get-PasswordHash

    return $computedHash -eq $StoredHash
}


# --- User Authentication for BROKER ---
function Authenticate-User {
    param(
        [Parameter(Mandatory=$true)][string]$Username,
        [Parameter(Mandatory=$true)][string]$PasswordPlainText,
        [Parameter(Mandatory=$true)][string]$UserAccountsFolderPath
    )
    $userFilePath = Join-Path -Path $UserAccountsFolderPath -ChildPath "$Username.txt"
    if (-not (Test-Path $userFilePath -PathType Leaf)) {
        Write-Warning "User account file not found: '$userFilePath'"
        return $null
    }

    $userData = @{}
    try {
        # Read lines, skipping blank/comment lines
        Get-Content -Path $userFilePath | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and -not $_.TrimStart().StartsWith('#') } | ForEach-Object {
            if ($_ -match '^\s*([^=]+?)\s*=\s*(.*?)\s*$') { # Match Key=Value, trimming whitespace
                $userData[$Matches[1]] = $Matches[2]
            }
        }
    } catch {
        Write-Error "Error reading user account file '$userFilePath': $($_.Exception.Message)"
        return $null
    }

    if (-not $userData.ContainsKey('Username') -or -not $userData.ContainsKey('PasswordHash')) {
        Write-Warning "User account file '$userFilePath' is missing 'Username' or 'PasswordHash'."
        return $null
    }

    # Optional: Check if Username in file matches requested Username
    # if ($userData.Username -ne $Username) {
    #     Write-Warning "Username mismatch in file '$userFilePath'."
    #     return $null
    # }

    if (Test-PasswordHash -PasswordPlainText $PasswordPlainText -StoredHash $userData.PasswordHash) {
        Write-Host "Authentication successful for user '$Username'." -ForegroundColor Green
        # Return a hashtable for consistency
        return [hashtable]@{
            Username = $userData.Username
            # Add other user properties from the file if needed, e.g., Role, FullName
        }
    } else {
        Write-Warning "Authentication failed for user '$Username'."
        return $null
    }
}

# --- Customer Profile Loading ---
# Loads ALL customer profiles from the specified folder
function Load-AllCustomerProfiles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$UserAccountsFolderPath # Path to the 'customer_accounts' folder
    )
    $allProfiles = @{}
    Write-Verbose "Loading all customer profiles from: $UserAccountsFolderPath"

    if (-not (Test-Path -Path $UserAccountsFolderPath -PathType Container)) {
        Write-Warning "Customer accounts folder '$UserAccountsFolderPath' not found."
        return $allProfiles
    }

    $profileFiles = Get-ChildItem -Path $UserAccountsFolderPath -Filter "*.txt" -File -ErrorAction SilentlyContinue
    if ($profileFiles) {
        foreach ($file in $profileFiles) {
            $customerNameFromFile = $file.BaseName
            Write-Verbose "Processing customer file '$($file.Name)'..."

            $profileData = [hashtable]@{
                "CustomerFileName"   = $customerNameFromFile
                "CustomerName"       = $customerNameFromFile # Default CustomerName to filename; can be overridden by file content
                "AllowedCentralKeys" = @()
                "AllowedSAIAKeys"    = @()
                "AllowedRLKeys"      = @()
                "AllowedAverittKeys" = @()
                "AllowedAAACooperKeys" = @() # CORRECTED: Added AllowedAAACooperKeys initialization
            }
            Write-Verbose "DEBUG Load-AllCustomerProfiles: Initialized profileData for '$customerNameFromFile'. Keys: $($profileData.Keys -join ', ')"

            try {
                # --- Line-by-line parsing ---
                Get-Content -Path $file.FullName -ErrorAction Stop | ForEach-Object {
                    $line = $_.Trim()
                    # Match Key=Value, trimming whitespace around key, value, and equals
                    if (-not ([string]::IsNullOrWhiteSpace($line)) -and $line -match '^\s*([^=]+?)\s*=\s*(.*?)\s*$') {
                        $key = $Matches[1]
                        $value = $Matches[2]

                        if ($key -eq 'CustomerName') { # Explicitly handle CustomerName if in file
                            if (-not [string]::IsNullOrWhiteSpace($value)) {
                                $profileData.CustomerName = $value # Override filename default
                                Write-Verbose "  Parsed CustomerName = '$value'"
                            }
                        } elseif ($key -like "Allowed*Keys") {
                            if (-not [string]::IsNullOrWhiteSpace($value)) {
                                # Split by comma, trim, and filter out any empty strings resulting from split (e.g. "val1,,val2")
                                $parsedArray = ($value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                                $profileData[$key] = $parsedArray
                                Write-Verbose "  Parsed $key = $($parsedArray -join '; ')"
                            } else {
                                # If line is "AllowedSomeKeys=", value is empty, so assign an empty array.
                                # This correctly overwrites the pre-initialized empty array with another empty array.
                                $profileData[$key] = @()
                                Write-Verbose "  Parsed $key = (empty value -> @())"
                            }
                        } else {
                            # For any other keys found in the file (e.g., PasswordHash)
                            $profileData[$key] = $value
                            Write-Verbose "  Parsed $key = '$value'"
                        }
                    } else {
                         Write-Verbose "  Skipping line (no match or blank): $line"
                    }
                } # End Get-Content | ForEach-Object

                # --- Post-parsing checks/cleanup (Optional but ensures array type) ---
                # CORRECTED: Added 'AllowedAAACooperKeys' to this loop
                foreach($keyTypeForArrayCheck in @('AllowedCentralKeys', 'AllowedSAIAKeys', 'AllowedRLKeys', 'AllowedAverittKeys', 'AllowedAAACooperKeys')) {
                    # Check if the key exists AND if its value is NOT already an array
                    if ($profileData.ContainsKey($keyTypeForArrayCheck) -and -not ($profileData[$keyTypeForArrayCheck] -is [array])) {
                        Write-Warning "Value for $keyTypeForArrayCheck in profile '$($profileData.CustomerName)' was not an array after parsing. Converting."
                        # If it was a single non-empty string, make it an array of one.
                        if ($profileData[$keyTypeForArrayCheck] -is [string] -and -not [string]::IsNullOrWhiteSpace($profileData[$keyTypeForArrayCheck])) {
                            $profileData[$keyTypeForArrayCheck] = @($profileData[$keyTypeForArrayCheck])
                        } else {
                            # Otherwise (e.g., empty string, $null, other type), force to empty array.
                            $profileData[$keyTypeForArrayCheck] = @()
                        }
                    }
                    # If key doesn't exist, pre-initialization (now including AllowedAAACooperKeys) ensures it is @()
                }

                # --- Store the profile ---
                $customerKeyForStorage = $profileData.CustomerName # Use the potentially overridden CustomerName
                Write-Verbose "DEBUG Load-AllCustomerProfiles: Storing profile for '$customerNameFromFile' under key '$customerKeyForStorage'."
                Write-Verbose "DEBUG Load-AllCustomerProfiles: Final keys in profileData for '$customerKeyForStorage': $($profileData.Keys -join ', ')"
                $allProfiles[$customerKeyForStorage] = $profileData
                Write-Verbose "Load successful for customer profile: $customerKeyForStorage"

            } catch {
                Write-Warning "Load-AllCustomerProfiles: Could not process customer profile file '$($file.Name)'. Error: $($_.Exception.Message)"
                Write-Warning "Load-AllCustomerProfiles: Profile data for '$customerNameFromFile' might be incomplete or not stored."
            }
        } # End foreach ($file in $profileFiles)

    } else {
        Write-Verbose "No .txt profile files found in '$UserAccountsFolderPath'."
    }
    Write-Host "Loaded $($allProfiles.Count) customer profile(s)." -ForegroundColor Gray
    return $allProfiles
}

Write-Verbose "TMS Authentication and Profile Functions loaded."