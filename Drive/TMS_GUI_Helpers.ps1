# TMS_GUI_Helpers.ps1
# Description: Contains helper functions specifically for the TMS GUI,
#              such as functions to populate UI elements. Includes AAA Cooper.
#              This file should be dot-sourced by TMS_GUI.ps1.

# Assumes $script scoped variables from TMS_GUI.ps1 (like $script:allCentralKeys, $script:allCustomerProfiles, etc.) are available.
# Assumes UI control variables (like $listBoxTariffs, $labelSelectedTariff) are in the caller's scope (TMS_GUI.ps1).
# Assumes TMS_Helpers_General.ps1 (with Get-PermittedKeys) is loaded.

# --- Helper Function to Populate Settings Tariff ListBox ---
Function Populate-TariffListBox {
    param(
        [Parameter(Mandatory)]$SelectedCarrier,
        [Parameter(Mandatory)]$ListBoxControl,
        [Parameter(Mandatory)]$LabelControl,
        [Parameter(Mandatory)]$ButtonControl,
        [Parameter(Mandatory)]$TextboxControl,
        [Parameter(Mandatory)][string]$SelectedCustomerName,
        [Parameter(Mandatory)][hashtable]$AllCustomerProfiles
    )
    $CustomerProfile = $null
    if (-not [string]::IsNullOrWhiteSpace($SelectedCustomerName) -and $AllCustomerProfiles.ContainsKey($SelectedCustomerName)) {
        $CustomerProfile = $AllCustomerProfiles[$SelectedCustomerName]
    }

    Write-Verbose "Populate-TariffListBox: Carrier='$SelectedCarrier', Customer='$SelectedCustomerName'"

    # Clear controls initially
    $ListBoxControl.BeginUpdate()
    $ListBoxControl.Items.Clear()
    $LabelControl.Text = "Selected: (None)"
    $ButtonControl.Enabled = $false
    $TextboxControl.Enabled = $false
    $TextboxControl.Clear()

    if ($null -eq $CustomerProfile) {
        Write-Warning "Populate-TariffListBox: Could not find profile data for '$SelectedCustomerName'."
        $ListBoxControl.Items.Add("Error: Customer '$SelectedCustomerName' not found.")
        $ListBoxControl.EndUpdate()
        return
    }

    # Determine which keys to load and which property holds allowed keys
    $allKeysForSelectedCarrier = $null
    $allowedKeyNamesProperty = $null
    switch ($SelectedCarrier) {
        "Central" { $allKeysForSelectedCarrier = $script:allCentralKeys; $allowedKeyNamesProperty = 'AllowedCentralKeys' }
        "SAIA"    { $allKeysForSelectedCarrier = $script:allSAIAKeys; $allowedKeyNamesProperty = 'AllowedSAIAKeys' }
        "RL"      { $allKeysForSelectedCarrier = $script:allRLKeys; $allowedKeyNamesProperty = 'AllowedRLKeys' }
        "Averitt" { $allKeysForSelectedCarrier = $script:allAverittKeys; $allowedKeyNamesProperty = 'AllowedAverittKeys' }
        "AAACooper" { $allKeysForSelectedCarrier = $script:allAAACooperKeys; $allowedKeyNamesProperty = 'AllowedAAACooperKeys' } # Added AAA Cooper
        default {
            Write-Warning "Populate-TariffListBox: Unknown carrier '$SelectedCarrier'."
            $ListBoxControl.Items.Add("Error: Unknown Carrier")
            $ListBoxControl.EndUpdate()
            return
        }
    }

    $allowedKeyNames = @()
    if ($CustomerProfile.ContainsKey($allowedKeyNamesProperty)) {
        $allowedKeyNames = $CustomerProfile[$allowedKeyNamesProperty]
        if ($null -eq $allowedKeyNames -or -not ($allowedKeyNames -is [array])) { $allowedKeyNames = @() }
    } else {
        Write-Warning "Populate-TariffListBox: Customer profile '$($CustomerProfile.CustomerName)' MISSING property '$allowedKeyNamesProperty'."
        $allowedKeyNames = @()
    }
    Write-Verbose "Populate-TariffListBox: Found $($allowedKeyNames.Count) allowed key names for $allowedKeyNamesProperty."

    $permittedKeys = @{}
    try {
        if ($null -ne $allKeysForSelectedCarrier) {
             if (Get-Command Get-PermittedKeys -ErrorAction SilentlyContinue) {
                 $permittedKeys = Get-PermittedKeys -AllKeys $allKeysForSelectedCarrier -AllowedKeyNames $allowedKeyNames
                 Write-Verbose "Populate-TariffListBox: Get-PermittedKeys returned $($permittedKeys.Count) items."
             } else {
                 Write-Error "Populate-TariffListBox: Get-PermittedKeys function not found."
                 $ListBoxControl.Items.Add("Error: Missing function.")
             }
        } else { Write-Warning "Populate-TariffListBox: Keys for carrier '$SelectedCarrier' not loaded (or $script:all...Keys variable is null)." }
    } catch {
        Write-Warning "Populate-TariffListBox: ERROR calling Get-PermittedKeys: $($_.Exception.Message)"
        $ListBoxControl.Items.Add("Error retrieving permissions.")
    }

    if ($permittedKeys.Count -gt 0) {
        $keyNamesSorted = @($permittedKeys.Keys | Sort-Object)
        foreach ($keyName in $keyNamesSorted) {
            $keyData = $permittedKeys[$keyName]
            $currentMargin = "N/A"
            if ($keyData -is [hashtable] -and $keyData.ContainsKey('MarginPercent')) {
                if ($keyData['MarginPercent'] -ne $null -and $keyData['MarginPercent'] -as [double] -ne $null) {
                     try {
                         $currentMargin = "{0:N1}%" -f ([double]$keyData['MarginPercent'])
                     } catch {
                          Write-Warning "Error formatting MarginPercent '$($keyData['MarginPercent'])' for key '$keyName': $($_.Exception.Message)"
                          $currentMargin = "FormatErr!"
                     }
                } else { $currentMargin = "Invalid!" }
            }
            $displayString = "{0,-40} {1,10}" -f $keyName, $currentMargin
            $ListBoxControl.Items.Add($displayString) | Out-Null
        }
    } else {
        Write-Warning "Populate-TariffListBox: No permitted keys found for '$SelectedCarrier' / '$($CustomerProfile.CustomerName)' after filtering."
        $ListBoxControl.Items.Add("No permitted tariffs for this carrier.") | Out-Null
    }
    $ListBoxControl.EndUpdate()
}


# --- Helper Function to Populate Report Tariff ListBox(es) ---
Function Populate-ReportTariffListBoxes {
    param(
        [Parameter(Mandatory)]$SelectedCarrier,
        [Parameter(Mandatory)]$ReportType,
        [Parameter(Mandatory)][string]$SelectedCustomerName,
        [Parameter(Mandatory)][hashtable]$AllCustomerProfiles,
        [Parameter(Mandatory)]$ListBox1,
        [Parameter(Mandatory)]$Label1,
        [Parameter(Mandatory)]$ListBox2,
        [Parameter(Mandatory)]$Label2
    )
    $CustomerProfile = $null
    if (-not [string]::IsNullOrWhiteSpace($SelectedCustomerName) -and $AllCustomerProfiles.ContainsKey($SelectedCustomerName)) {
        $CustomerProfile = $AllCustomerProfiles[$SelectedCustomerName]
    }

    Write-Verbose "Populate-ReportTariffListBoxes: Carrier='$SelectedCarrier', Report='$ReportType', Customer='$SelectedCustomerName'"
    $ListBox1.BeginUpdate(); $ListBox1.Items.Clear(); $ListBox2.BeginUpdate(); $ListBox2.Items.Clear()

    $needsTwoLists = ($ReportType -eq "Carrier Comparison" -or $ReportType -eq "Avg Required Margin")
    $ListBox2.Visible = $needsTwoLists
    $Label2.Visible = $needsTwoLists
    if ($needsTwoLists) { $Label1.Text = "Tariff 1 (Base):" } else { $Label1.Text = "Select Tariff:" }

    if ($null -eq $CustomerProfile) {
        Write-Warning "Populate-ReportTariffListBoxes: Could not find profile data for '$SelectedCustomerName'."
        $ListBox1.Items.Add("Select Customer"); $ListBox1.EndUpdate(); $ListBox2.Items.Add("Select Customer"); $ListBox2.EndUpdate(); return
    }

    $allKeysForSelectedCarrier = $null; $allowedKeyNamesProperty = $null
    switch ($SelectedCarrier) {
        "Central" { $allKeysForSelectedCarrier = $script:allCentralKeys; $allowedKeyNamesProperty = 'AllowedCentralKeys' }
        "SAIA"    { $allKeysForSelectedCarrier = $script:allSAIAKeys; $allowedKeyNamesProperty = 'AllowedSAIAKeys' }
        "RL"      { $allKeysForSelectedCarrier = $script:allRLKeys; $allowedKeyNamesProperty = 'AllowedRLKeys' }
        "Averitt" { $allKeysForSelectedCarrier = $script:allAverittKeys; $allowedKeyNamesProperty = 'AllowedAverittKeys' }
        "AAACooper" { $allKeysForSelectedCarrier = $script:allAAACooperKeys; $allowedKeyNamesProperty = 'AllowedAAACooperKeys' } # Added AAA Cooper
        default   { Write-Warning "Populate-ReportTariffListBoxes: Unknown carrier '$SelectedCarrier'."; $ListBox1.Items.Add("Unknown Carrier"); $ListBox1.EndUpdate(); $ListBox2.Items.Add("Unknown Carrier"); $ListBox2.EndUpdate(); return }
    }

    $allowedKeyNames = @()
    if ($CustomerProfile.ContainsKey($allowedKeyNamesProperty)) {
        $allowedKeyNames = $CustomerProfile[$allowedKeyNamesProperty]
        if ($null -eq $allowedKeyNames -or -not ($allowedKeyNames -is [array])) { $allowedKeyNames = @() }
    } else {
        Write-Warning "Populate-ReportTariffListBoxes: Customer profile '$($CustomerProfile.CustomerName)' MISSING property '$allowedKeyNamesProperty'."
        $allowedKeyNames = @()
    }
    Write-Verbose "Populate-ReportTariffListBoxes: Found $($allowedKeyNames.Count) allowed key names for $allowedKeyNamesProperty."

    $permittedKeys = @{}
    try {
        if ($null -ne $allKeysForSelectedCarrier) {
             if (Get-Command Get-PermittedKeys -ErrorAction SilentlyContinue) {
                $permittedKeys = Get-PermittedKeys -AllKeys $allKeysForSelectedCarrier -AllowedKeyNames $allowedKeyNames
                Write-Verbose "Populate-ReportTariffListBoxes: Get-PermittedKeys returned $($permittedKeys.Count) items."
             } else { Write-Error "Populate-ReportTariffListBoxes: Get-PermittedKeys function not found."; $ListBox1.Items.Add("Error") }
        } else { Write-Warning "Populate-ReportTariffListBoxes: Keys for carrier '$SelectedCarrier' not loaded (or $script:all...Keys variable is null)." }
    } catch {
        Write-Warning "Populate-ReportTariffListBoxes: Error getting permitted keys: $($_.Exception.Message)"
        $ListBox1.Items.Add("Error");
    }

    if ($permittedKeys.Count -gt 0) {
        $keyNamesSorted = @($permittedKeys.Keys | Sort-Object)
        $ListBox1.Items.AddRange($keyNamesSorted)
        if ($needsTwoLists) { $ListBox2.Items.AddRange($keyNamesSorted) }
    } else {
        Write-Warning "Populate-ReportTariffListBoxes: No permitted tariffs for '$SelectedCarrier' / '$($CustomerProfile.CustomerName)'."
        $ListBox1.Items.Add("No permitted tariffs")
        if ($needsTwoLists) { $ListBox2.Items.Add("No permitted tariffs") }
    }
    $ListBox1.EndUpdate(); $ListBox2.EndUpdate()
}

Write-Verbose "TMS GUI Helper Functions loaded."
