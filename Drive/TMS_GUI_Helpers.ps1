# TMS_GUI_Helpers.ps1
# Description: Contains helper functions specifically for the TMS GUI,
#              such as functions to populate UI elements.
#              This file should be dot-sourced by TMS_GUI.ps1.

# Assumes $script scoped variables from TMS_GUI.ps1 (like $script:allCentralKeys, $script:selectedCustomerProfile, etc.) are available.
# Assumes UI control variables (like $listBoxTariffs, $labelSelectedTariff) are in the caller's scope (TMS_GUI.ps1).

# --- Helper Function to Populate Settings Tariff ListBox ---
Function Populate-TariffListBox { 
    param(
        [Parameter(Mandatory)]$SelectedCarrier, 
        [Parameter(Mandatory)]$ListBoxControl, 
        [Parameter(Mandatory)]$LabelControl, 
        [Parameter(Mandatory)]$ButtonControl, 
        [Parameter(Mandatory)]$TextboxControl, 
        [Parameter(Mandatory)]$CustomerProfile # This is $script:selectedCustomerProfile from TMS_GUI.ps1
    )
    Write-Host "DEBUG GUI_Helpers (Populate-TariffListBox): Function called. Carrier: '$SelectedCarrier', Customer: '$($CustomerProfile.CustomerName)'"
    
    if ($null -eq $CustomerProfile) { 
        Write-Warning "DEBUG GUI_Helpers (Populate-TariffListBox): CustomerProfile is NULL! Cannot populate."
        $ListBoxControl.Items.Clear()
        $ListBoxControl.Items.Add("Error: No Customer Selected.")
        # Potentially update a status bar if available in the main GUI scope
        # if ($script:statusBar) { $script:statusBar.Text = "Error: No customer selected for settings." }
        return 
    }

    $ListBoxControl.BeginUpdate()
    $ListBoxControl.Items.Clear()
    $LabelControl.Text = "Selected: (None)"
    $ButtonControl.Enabled = $false
    $TextboxControl.Enabled = $false
    $TextboxControl.Clear()

    $allKeysForSelectedCarrier = $null
    $allowedKeyNamesProperty = $null

    switch ($SelectedCarrier) {
        "Central" { 
            $allKeysForSelectedCarrier = $script:allCentralKeys # Accessing script-scoped variable
            $allowedKeyNamesProperty = 'AllowedCentralKeys'
        }
        "SAIA"    { 
            $allKeysForSelectedCarrier = $script:allSAIAKeys # Accessing script-scoped variable
            $allowedKeyNamesProperty = 'AllowedSAIAKeys'
        }
        "RL"      { 
            $allKeysForSelectedCarrier = $script:allRLKeys # Accessing script-scoped variable
            $allowedKeyNamesProperty = 'AllowedRLKeys'
        }
        default { 
            Write-Warning "DEBUG GUI_Helpers (Populate-TariffListBox): Unknown carrier '$SelectedCarrier'."
            $ListBoxControl.Items.Add("Error: Unknown Carrier")
            $ListBoxControl.EndUpdate()
            return 
        }
    }
    Write-Host "DEBUG GUI_Helpers (Populate-TariffListBox): Selected AllKeys for '$SelectedCarrier'. Count: $(if($null -ne $allKeysForSelectedCarrier){$allKeysForSelectedCarrier.Count}else{'NULL or 0'}). IsHashtable: $($allKeysForSelectedCarrier -is [hashtable])"

    $allowedKeyNames = @()
    if ($CustomerProfile.PSObject.Properties.Name -contains $allowedKeyNamesProperty) {
        $allowedKeyNames = $CustomerProfile.$allowedKeyNamesProperty
        if ($null -eq $allowedKeyNames) { $allowedKeyNames = @() } 
        Write-Host "DEBUG GUI_Helpers (Populate-TariffListBox): AllowedKeyNames for '$($allowedKeyNamesProperty)': $($allowedKeyNames -join ', ' | Out-String). Count: $($allowedKeyNames.Count)"
    } else {
        Write-Warning "DEBUG GUI_Helpers (Populate-TariffListBox): Customer profile for '$($CustomerProfile.CustomerName)' is missing property '$allowedKeyNamesProperty'."
    }

    $permittedKeys = @{}
    try {
        if (-not (Get-Command Get-PermittedKeys -ErrorAction SilentlyContinue)) { 
            throw "Get-PermittedKeys function (from TMS_Helpers.ps1) not found." 
        }
        if ($null -eq $allKeysForSelectedCarrier) {
             Write-Warning "DEBUG GUI_Helpers (Populate-TariffListBox): AllKeysForSelectedCarrier is NULL for '$SelectedCarrier'. Cannot call Get-PermittedKeys."
        } elseif($allKeysForSelectedCarrier.Count -eq 0) {
            Write-Warning "DEBUG GUI_Helpers (Populate-TariffListBox): AllKeysForSelectedCarrier is EMPTY for '$SelectedCarrier'. Get-PermittedKeys will likely return empty."
            # No need to call Get-PermittedKeys if there are no keys to filter from
        }
         else {
            Write-Host "DEBUG GUI_Helpers (Populate-TariffListBox): Calling Get-PermittedKeys with AllKeys count: $($allKeysForSelectedCarrier.Count), AllowedKeyNames count: $($allowedKeyNames.Count)"
            $permittedKeys = Get-PermittedKeys -AllKeys $allKeysForSelectedCarrier -AllowedKeyNames $allowedKeyNames
            Write-Host "DEBUG GUI_Helpers (Populate-TariffListBox): Get-PermittedKeys returned. Count = $($permittedKeys.Count). Keys: $($permittedKeys.Keys -join ', ' | Out-String)"
        }
    } catch {
        Write-Warning "DEBUG GUI_Helpers (Populate-TariffListBox): ERROR calling Get-PermittedKeys: $($_.Exception.Message)"
        # if ($script:statusBar) { $script:statusBar.Text = "Error getting permitted keys for ${SelectedCarrier}: ${$_.Exception.Message}" }
        $ListBoxControl.Items.Add("Error retrieving permissions.")
        $ListBoxControl.EndUpdate()
        return
    }

    if ($permittedKeys.Count -gt 0) {
        Write-Host "DEBUG GUI_Helpers (Populate-TariffListBox): Populating listbox with $($permittedKeys.Count) items..."
        $keyNamesSorted = @($permittedKeys.Keys | Sort-Object)
        foreach ($keyName in $keyNamesSorted) {
            $keyData = $permittedKeys[$keyName]
            $currentMargin = "N/A"
            if ($keyData -is [hashtable] -and $keyData.ContainsKey('MarginPercent')) {
                if ($keyData['MarginPercent'] -as [double] -ne $null) {
                    $currentMargin = "{0:N1}%" -f ([double]$keyData['MarginPercent'])
                } else { $currentMargin = "Invalid!" }
            }
            $displayString = "{0,-30} {1,10}" -f $keyName, $currentMargin
            $ListBoxControl.Items.Add($displayString) | Out-Null
        }
    } else {
        Write-Warning "DEBUG GUI_Helpers (Populate-TariffListBox): No permitted keys found for '$SelectedCarrier' and customer '$($CustomerProfile.CustomerName)' after filtering."
        $ListBoxControl.Items.Add("No permitted tariffs for this carrier.") | Out-Null
    }
    $ListBoxControl.EndUpdate()
}


# --- Helper Function to Populate Report Tariff ListBox(es) ---
Function Populate-ReportTariffListBoxes { 
    param(
        [Parameter(Mandatory)]$SelectedCarrier, 
        [Parameter(Mandatory)]$ReportType, 
        [Parameter(Mandatory)]$CustomerProfile, # This is $script:selectedCustomerProfile from TMS_GUI.ps1
        [Parameter(Mandatory)]$ListBox1, 
        [Parameter(Mandatory)]$Label1, 
        [Parameter(Mandatory)]$ListBox2, 
        [Parameter(Mandatory)]$Label2
    )
    Write-Host "DEBUG GUI_Helpers (Populate-ReportTariffListBoxes): Carrier='$SelectedCarrier', Report='$ReportType', Customer='$($CustomerProfile.CustomerName)'"
    $ListBox1.BeginUpdate(); $ListBox1.Items.Clear(); $ListBox2.BeginUpdate(); $ListBox2.Items.Clear()

    $needsTwoLists = ($ReportType -eq "Carrier Comparison" -or $ReportType -eq "Avg Required Margin")
    $ListBox2.Visible = $needsTwoLists
    $Label2.Visible = $needsTwoLists
    if ($needsTwoLists) { $Label1.Text = "Tariff 1 (Base):" } else { $Label1.Text = "Select Tariff:" }

    if ($null -eq $CustomerProfile) { 
        Write-Warning "DEBUG GUI_Helpers (Populate-ReportTariffListBoxes): CustomerProfile is NULL."
        $ListBox1.Items.Add("Select Customer"); $ListBox1.EndUpdate(); $ListBox2.EndUpdate(); return 
    }

    $allKeysForSelectedCarrier = $null; $allowedKeyNamesProperty = $null
    switch ($SelectedCarrier) {
        "Central" { $allKeysForSelectedCarrier = $script:allCentralKeys; $allowedKeyNamesProperty = 'AllowedCentralKeys' }
        "SAIA"    { $allKeysForSelectedCarrier = $script:allSAIAKeys; $allowedKeyNamesProperty = 'AllowedSAIAKeys' }
        "RL"      { $allKeysForSelectedCarrier = $script:allRLKeys; $allowedKeyNamesProperty = 'AllowedRLKeys' }
        default   { Write-Warning "DEBUG GUI_Helpers (Populate-ReportTariffListBoxes): Unknown carrier '$SelectedCarrier'."; $ListBox1.EndUpdate(); $ListBox2.EndUpdate(); return }
    }
    Write-Host "DEBUG GUI_Helpers (Populate-ReportTariffListBoxes): AllKeys for '$SelectedCarrier' Count: $(if($null -ne $allKeysForSelectedCarrier){$allKeysForSelectedCarrier.Count}else{'NULL or 0'})"


    $allowedKeyNames = @()
    if ($CustomerProfile.PSObject.Properties.Name -contains $allowedKeyNamesProperty) {
        $allowedKeyNames = $CustomerProfile.$allowedKeyNamesProperty
        if ($null -eq $allowedKeyNames) { $allowedKeyNames = @() }
         Write-Host "DEBUG GUI_Helpers (Populate-ReportTariffListBoxes): AllowedKeyNames for '$($allowedKeyNamesProperty)': $($allowedKeyNames -join ', ' | Out-String). Count: $($allowedKeyNames.Count)"
    } else {
         Write-Warning "DEBUG GUI_Helpers (Populate-ReportTariffListBoxes): Customer profile for '$($CustomerProfile.CustomerName)' is missing property '$allowedKeyNamesProperty'."
    }
    
    $permittedKeys = @{}
    try {
        if ($null -eq $allKeysForSelectedCarrier) {
            Write-Warning "DEBUG GUI_Helpers (Populate-ReportTariffListBoxes): AllKeysForSelectedCarrier is NULL for '$SelectedCarrier'. Cannot call Get-PermittedKeys."
        } elseif($allKeysForSelectedCarrier.Count -eq 0) {
             Write-Warning "DEBUG GUI_Helpers (Populate-ReportTariffListBoxes): AllKeysForSelectedCarrier is EMPTY for '$SelectedCarrier'. Get-PermittedKeys will likely return empty."
        } else {
            $permittedKeys = Get-PermittedKeys -AllKeys $allKeysForSelectedCarrier -AllowedKeyNames $allowedKeyNames
             Write-Host "DEBUG GUI_Helpers (Populate-ReportTariffListBoxes): Get-PermittedKeys returned. Count = $($permittedKeys.Count). Keys: $($permittedKeys.Keys -join ', ' | Out-String)"
        }
    } catch { 
        Write-Warning "DEBUG GUI_Helpers (Populate-ReportTariffListBoxes): Error getting permitted keys for report: $($_.Exception.Message)"
        $ListBox1.Items.Add("Error"); $ListBox1.EndUpdate(); $ListBox2.EndUpdate(); return 
    }

    if ($permittedKeys.Count -gt 0) {
        $keyNamesSorted = @($permittedKeys.Keys | Sort-Object)
        $ListBox1.Items.AddRange($keyNamesSorted)
        if ($needsTwoLists) { $ListBox2.Items.AddRange($keyNamesSorted) }
    } else {
        Write-Warning "DEBUG GUI_Helpers (Populate-ReportTariffListBoxes): No permitted tariffs for '$SelectedCarrier' and customer '$($CustomerProfile.CustomerName)'."
        $ListBox1.Items.Add("No permitted tariffs")
        if ($needsTwoLists) { $ListBox2.Items.Add("No permitted tariffs") }
    }
    $ListBox1.EndUpdate(); $ListBox2.EndUpdate()
}

Write-Verbose "TMS GUI Helper Functions loaded."
