# TMS_Auth.ps1
# Description: Contains functions for user authentication (for brokers)
#              and loading customer profiles.
#              This file should be dot-sourced by the main entry script (TMS_GUI.ps1).

# --- Password Hashing Functions ---
# IMPORTANT: These are simplified examples using basic SHA256 without salt.
# For production systems, use a robust, salted hashing algorithm.

function New-PasswordHash {
    # Creates a SHA256 hash of a plain text password.
    param(
        [Parameter(Mandatory=$true)]
        [string]$PlainTextPassword
    )
    try {
        $sha256 = [System.Security.Cryptography.SHA256]::Create()
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($PlainTextPassword)
        $hashBytes = $sha256.ComputeHash($bytes)
        # Convert hash bytes to a hex string for storage
        $hashString = ($hashBytes | ForEach-Object { $_.ToString('x2') }) -join ''
        $sha256.Dispose() # Dispose the crypto object
        return $hashString
    } catch {
        Write-Error "Failed to generate password hash: $($_.Exception.Message)"
        return $null
    }
}

function Test-PasswordHash {
    # Compares a plain text password against a stored SHA256 hash.
     param(
        [Parameter(Mandatory=$true)]
        [string]$PlainTextPassword,
        [Parameter(Mandatory=$true)]
        [string]$StoredHash
    )
    try {
        # Ensure New-PasswordHash function exists before calling
        if (-not (Get-Command New-PasswordHash -ErrorAction SilentlyContinue)) {
             Write-Error "FATAL: New-PasswordHash function not found within Test-PasswordHash. Check TMS_Auth.ps1 structure."
             throw "New-PasswordHash function not found."
        }
        Write-Host "DEBUG (Test-PasswordHash): Hashing input password." # DEBUG
        $newHash = New-PasswordHash -PlainTextPassword $PlainTextPassword
        Write-Host "DEBUG (Test-PasswordHash): InputHash: '$newHash', StoredHash: '$StoredHash'" # DEBUG
        return $newHash -eq $StoredHash # Case-insensitive comparison is default for -eq with strings
    } catch {
         Write-Error "Failed to test password hash: $($_.Exception.Message)"
         return $false
    }
}

# --- Broker Authentication Functions ---

function Get-UserCredentials {
    # Prompts the user for username and password for broker login.
    # Returns a PSCustomObject with Username and PasswordPlainText or $null if cancelled.
    # This function is typically called by the GUI's login window.
    Write-Host "`nPlease Login (Broker Account)" -ForegroundColor Yellow
    $username = Read-Host " Username"
    if ([string]::IsNullOrWhiteSpace($username)) { Write-Warning "Username cannot be empty."; return $null }
    $password = Read-Host " Password" -AsSecureString
    if ($password -eq $null -or $password.Length -eq 0) { Write-Warning "Password cannot be empty."; return $null }

    # Convert SecureString to plain text *only for immediate use* in Authenticate-User
    $credential = New-Object System.Management.Automation.PSCredential($username, $password)
    $plainTextPassword = $credential.GetNetworkCredential().Password

    return [PSCustomObject]@{
        Username = $username
        PasswordPlainText = $plainTextPassword
    }
}

function Authenticate-User {
    # Authenticates a BROKER user based on username and plain text password against stored profiles.
    # Returns the user profile hashtable on success, $null on failure.
    param(
        [Parameter(Mandatory=$true)] [string]$Username,
        [Parameter(Mandatory=$true)] [string]$PasswordPlainText,
        [Parameter(Mandatory=$true)] [string]$UserAccountsFolderPath # Path to BROKER account files (e.g., user_accounts/)
    )
    Write-Host "DEBUG (Authenticate-User): Starting authentication for '$Username'." 
    Write-Host "DEBUG (Authenticate-User): UserAccountsFolderPath is '$UserAccountsFolderPath'." 

    $userFilePath = Join-Path -Path $UserAccountsFolderPath -ChildPath "$($Username).txt"
    Write-Host "DEBUG (Authenticate-User): Checking for user file at path: '$userFilePath'" 

    if (-not (Test-Path $userFilePath -PathType Leaf)) {
        Write-Warning "Broker account file not found for '$Username' at '$userFilePath'." 
        Write-Host "DEBUG (Authenticate-User): User file not found. Returning null." 
        return $null
    }

    $userProfile = @{ Username = $Username } 
    try {
        Write-Host "DEBUG (Authenticate-User): Reading user file '$userFilePath'." 
        $lines = Get-Content -Path $userFilePath -ErrorAction Stop
        foreach ($line in $lines) {
            $trimmedLine = $line.Trim()
            if ([string]::IsNullOrWhiteSpace($trimmedLine) -or $trimmedLine.StartsWith('#')) { continue }

            $equalsIndex = $trimmedLine.IndexOf('=')
            if ($equalsIndex -gt 0) {
                $key = $trimmedLine.Substring(0, $equalsIndex).Trim()
                $value = $trimmedLine.Substring($equalsIndex + 1).Trim()

                if (-not [string]::IsNullOrEmpty($key)) {
                    Write-Verbose "DEBUG (Authenticate-User Parse): Line='$trimmedLine', Key='$key', Value='$value'" 
                    # For broker profiles, we primarily care about PasswordHash and potentially other direct properties.
                    # Allowed*Keys are typically for customer profiles.
                    if ($key -notlike "Allowed*Keys") {
                        Write-Verbose "DEBUG (Authenticate-User Parse): Storing property for Broker Key='$key'" 
                        $userProfile[$key] = $value
                    } else {
                         Write-Verbose "DEBUG (Authenticate-User Parse): Skipping Allowed*Keys for Broker Key='$key'" 
                    }
                }
            } else {
                Write-Warning "Skipping line (no '=') in broker profile '$($Username).txt': $line"
            }
        } 
        Write-Host "DEBUG (Authenticate-User): Finished parsing profile. Keys loaded: $($userProfile.Keys -join ', ')" 

        if (-not $userProfile.ContainsKey('PasswordHash')) {
            Write-Warning "Broker profile for '$Username' is missing 'PasswordHash'. Cannot authenticate."
            Write-Host "DEBUG (Authenticate-User): PasswordHash missing. Returning null." 
            return $null
        }

        $storedHash = $userProfile.PasswordHash
        Write-Host "DEBUG (Authenticate-User): StoredHash retrieved: '$storedHash' (length: $($storedHash.Length))" 

        if (-not (Get-Command Test-PasswordHash -ErrorAction SilentlyContinue)) {
             Write-Error "FATAL: Test-PasswordHash function not found within Authenticate-User."
             throw "Test-PasswordHash function not found."
        }

        Write-Host "DEBUG (Authenticate-User): Calling Test-PasswordHash." 
        $isMatch = Test-PasswordHash -PlainTextPassword $PasswordPlainText -StoredHash $storedHash
        Write-Host "DEBUG (Authenticate-User): Test-PasswordHash result: $isMatch" 

        if ($isMatch) {
             Write-Verbose "Password verified for broker user '$Username'."
             Write-Host "DEBUG (Authenticate-User): Password MATCH. Returning profile." 
             return $userProfile
        } else {
             Write-Warning "Incorrect password for broker user '$Username'."
             Write-Host "DEBUG (Authenticate-User): Password MISMATCH. Returning null." 
             return $null
        }

    } catch {
        Write-Error "Failed to load or process broker profile '$($Username).txt': $($_.Exception.Message)"
        Write-Host "DEBUG (Authenticate-User): Exception during profile processing. Returning null. Error: $($_.Exception.Message)" 
        return $null
    }
}

# --- Customer Profile Loading Function ---
function Load-AllCustomerProfiles {
    # Loads all customer profiles from the specified customer accounts folder.
    # Customer profiles define permissions like AllowedCarrierKeys.
    param(
        [Parameter(Mandatory=$true)]
        [string]$CustomerAccountsFolderPath # Path to CUSTOMER account files (e.g., customer_accounts/)
    )
    $customerProfiles = @{}
    Write-Verbose "Loading all customer profiles from: $CustomerAccountsFolderPath"
    if (-not (Test-Path -Path $CustomerAccountsFolderPath -PathType Container)) {
        Write-Warning "Customer accounts folder not found at '$CustomerAccountsFolderPath'. Cannot load customer profiles."
        return $customerProfiles # Return empty hashtable
    }

    $profileFiles = Get-ChildItem -Path $CustomerAccountsFolderPath -Filter "*.txt" -File -ErrorAction SilentlyContinue
    if ($profileFiles) {
        foreach ($file in $profileFiles) {
            $customerName = $file.BaseName
            $profileData = @{ CustomerName = $customerName } # Initialize with CustomerName
            try {
                $lines = Get-Content -Path $file.FullName -ErrorAction Stop
                foreach ($line in $lines) {
                    $trimmedLine = $line.Trim()
                    if ([string]::IsNullOrWhiteSpace($trimmedLine) -or $trimmedLine.StartsWith('#')) { continue }
                    $equalsIndex = $trimmedLine.IndexOf('=')
                    if ($equalsIndex -gt 0) {
                        $key = $trimmedLine.Substring(0, $equalsIndex).Trim()
                        $value = $trimmedLine.Substring($equalsIndex + 1).Trim()
                        if (-not [string]::IsNullOrEmpty($key)) {
                            if ($key -like "Allowed*Keys") {
                                 # Split comma-separated values into an array, trim whitespace from each, and filter out empty strings
                                 $profileData[$key] = @($value.Split(',') | ForEach-Object { $_.Trim() } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                            } elseif ($key -ne 'PasswordHash') { # Exclude password hash from customer profile data
                                $profileData[$key] = $value
                            }
                        }
                    }
                }

                # Ensure AllowedKeys arrays exist for all currently supported carriers
                if (-not $profileData.ContainsKey('AllowedCentralKeys')) { $profileData['AllowedCentralKeys'] = @() }
                if (-not $profileData.ContainsKey('AllowedSAIAKeys')) { $profileData['AllowedSAIAKeys'] = @() }
                if (-not $profileData.ContainsKey('AllowedRLKeys')) { $profileData['AllowedRLKeys'] = @() }
                if (-not $profileData.ContainsKey('AllowedAverittKeys')) { $profileData['AllowedAverittKeys'] = @() } # <<< AVERITT ADDED >>>
                
                $customerProfiles[$customerName] = $profileData
            } catch {
                Write-Warning "Could not process customer profile file '$($file.Name)': $($_.Exception.Message)"
            }
        }
    } else {
        Write-Warning "No customer profile files (.txt) found in '$CustomerAccountsFolderPath'."
    }
    Write-Host "Loaded $($customerProfiles.Count) customer profile(s)." -ForegroundColor Gray
    return $customerProfiles
}

# --- Helper/Admin Function (Example for managing broker passwords) ---
function Set-UserPassword {
    # Example function to generate a hash and update a user file.
    # Requires manual execution or integration into an admin tool.
    param(
        [Parameter(Mandatory=$true)] [string]$Username,
        [Parameter(Mandatory=$true)] [string]$UserAccountsFolderPath # Path to BROKER accounts
    )
    $userFilePath = Join-Path -Path $UserAccountsFolderPath -ChildPath "$($Username).txt"
    if (-not (Test-Path $userFilePath -PathType Leaf)) { Write-Error "User file '$userFilePath' not found."; return }

    $password = Read-Host "Enter NEW password for '$Username'" -AsSecureString
    if ($password -eq $null -or $password.Length -eq 0) { Write-Warning "Password cannot be empty."; return }

    $credential = New-Object System.Management.Automation.PSCredential($Username, $password)
    $plainTextPassword = $credential.GetNetworkCredential().Password
    $newHash = New-PasswordHash -PlainTextPassword $plainTextPassword
    if (-not $newHash) { Write-Error "Failed to generate hash. Password not updated."; return }

    try {
        $content = Get-Content -Path $userFilePath -Raw -ErrorAction Stop
        if ($content -match '(?m)^PasswordHash=.*') {
            $newContent = $content -replace "(?m)^PasswordHash=.*", "PasswordHash=$newHash"
        } else {
            $newContent = $content.TrimEnd() + [System.Environment]::NewLine + "PasswordHash=$newHash" # Ensure it's on a new line
        }
        Set-Content -Path $userFilePath -Value $newContent -Encoding UTF8 -ErrorAction Stop
        Write-Host "Password hash updated successfully for '$Username'." -ForegroundColor Green
    } catch {
        Write-Error "Failed to update password hash in '$userFilePath': $($_.Exception.Message)"
    }
}

Write-Verbose "TMS Authentication Functions loaded."
