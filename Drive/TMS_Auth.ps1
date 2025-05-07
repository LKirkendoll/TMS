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
        Get-Content -Path $userFilePath | ForEach-Object {
            $line = $_.Trim()
            if (-not ([string]::IsNullOrWhiteSpace($line)) -and $line -match '^(.+?)\s*=\s*(.*)$') {
                $userData[$Matches[1].Trim()] = $Matches[2].Trim()
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

    if ($userData.Username -ne $Username) {
        Write-Warning "Username mismatch in file '$userFilePath'." 
        return $null
    }

    if (Test-PasswordHash -PasswordPlainText $PasswordPlainText -StoredHash $userData.PasswordHash) {
        Write-Host "Authentication successful for user '$Username'." -ForegroundColor Green
        # Return a hashtable for consistency with other data structures if needed,
        # or PSCustomObject if preferred for user profiles.
        # For report functions expecting [hashtable], this should also be a hashtable.
        return [hashtable]@{ # <<< MODIFICATION: Ensure BrokerProfile is also a hashtable
            Username = $userData.Username
            # Add other user properties from the file if needed, e.g., Role, FullName
            # FullName = $userData.FullName 
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
            # <<< MODIFICATION: Initialize as Hashtable >>>
            $profileData = [hashtable]@{ "CustomerFileName" = $customerNameFromFile } 

            try {
                Get-Content -Path $file.FullName -ErrorAction Stop | ForEach-Object {
                    $line = $_.Trim()
                    if (-not ([string]::IsNullOrWhiteSpace($line)) -and $line -match '^(.+?)\s*=\s*(.*)$') {
                        $key = $Matches[1].Trim()
                        $value = $Matches[2].Trim()

                        if ($key -like "Allowed*Keys" -and $value -match ',') {
                            $profileData[$key] = ($value -split ',' | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                        } else {
                            $profileData[$key] = $value
                        }
                    }
                }

                if (-not $profileData.ContainsKey('CustomerName') -or [string]::IsNullOrWhiteSpace($profileData.CustomerName) ) {
                    $profileData['CustomerName'] = $customerNameFromFile
                }
                
                foreach($keyType in @('AllowedCentralKeys', 'AllowedSAIAKeys', 'AllowedRLKeys', 'AllowedAverittKeys')) { # Added Averitt for completeness from error
                    if ($profileData.ContainsKey($keyType) -and -not ($profileData[$keyType] -is [array])) {
                        # If it's a single string (not null/empty), convert to an array of one
                        if(-not [string]::IsNullOrWhiteSpace($profileData[$keyType])) {
                            $profileData[$keyType] = @($profileData[$keyType])
                        } else { # If it's an empty string or $null, make it an empty array
                            $profileData[$keyType] = @()
                        }
                    } elseif (-not $profileData.ContainsKey($keyType)) {
                        $profileData[$keyType] = @() 
                    }
                }
                # <<< MODIFICATION: No need to cast to PSCustomObject here, it's already a hashtable >>>
                $allProfiles[$profileData.CustomerName] = $profileData 
                Write-Verbose "Loaded customer profile: $($profileData.CustomerName)"

            } catch {
                Write-Warning "Could not process customer profile file '$($file.Name)': $($_.Exception.Message)"
            }
        }
    } else {
        Write-Verbose "No .txt profile files found in '$UserAccountsFolderPath'."
    }
    Write-Host "Loaded $($allProfiles.Count) customer profile(s)." -ForegroundColor Gray
    return $allProfiles
}

Write-Verbose "TMS Authentication and Profile Functions loaded."
