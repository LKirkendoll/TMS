<#
.SYNOPSIS
GUI Front-end for the Transportation Management System Tool.
Uses Windows Forms and leverages existing TMS PowerShell modules.
Broker logs in (from user_accounts), then selects a customer (from customer_accounts) to work with.
Includes Quote, Settings, and Reports tabs.
#>

# --- Load Required Assemblies ---
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
} catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to load required .NET Assemblies (System.Windows.Forms, System.Drawing). Ensure .NET Framework is available.", "Fatal Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    Exit 1
}

# --- Determine Script Root ---
$script:scriptRoot = $null
if ($PSScriptRoot) { $script:scriptRoot = $PSScriptRoot }
elseif ($MyInvocation.MyCommand.Path) { $script:scriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent }
else {
    [System.Windows.Forms.MessageBox]::Show("FATAL: Cannot determine script root directory.", "Fatal Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    Exit 1
}
Write-Verbose "Script Root: '$script:scriptRoot'"

# --- Dot-Source Module Files ---
Write-Verbose "Loading TMS Modules..."
try {
    # <<< MODIFICATION: Updated module list for split helpers >>>
    $moduleFiles = @(
        "TMS_Config.ps1",
        "TMS_Helpers_General.ps1",   # General utilities
        "TMS_GUI_Helpers.ps1",       # GUI-specific helpers
        "TMS_Auth.ps1",
        "TMS_Helpers_Central.ps1",   # Central-specific helpers
        "TMS_Helpers_SAIA.ps1",      # SAIA-specific helpers
        "TMS_Helpers_RL.ps1",        # R+L-specific helpers
        "TMS_Carrier_Central.ps1", 
        "TMS_Carrier_SAIA.ps1",    
        "TMS_Carrier_RL.ps1",      
        "TMS_Reports.ps1",
        "TMS_Settings.ps1"         
    )
    # <<< END MODIFICATION >>>

    foreach ($moduleFile in $moduleFiles) {
        $modulePath = Join-Path $script:scriptRoot $moduleFile
        Write-Host "DEBUG: Attempting to load module '$moduleFile' from '$modulePath'"
        if (-not (Test-Path $modulePath -PathType Leaf)) { throw "Module not found: '$moduleFile' at '$modulePath'" }

        try {
            . $modulePath
            Write-Host "DEBUG: Successfully dot-sourced '$moduleFile'."
            # --- Check for required functions immediately after loading relevant modules ---
            if ($moduleFile -eq "TMS_Auth.ps1") {
                if (-not (Get-Command Test-PasswordHash -ErrorAction SilentlyContinue)) { throw "Required function 'Test-PasswordHash' not found immediately after loading TMS_Auth.ps1." }
                if (-not (Get-Command Load-AllCustomerProfiles -ErrorAction SilentlyContinue)) { throw "Required function 'Load-AllCustomerProfiles' not found immediately after loading TMS_Auth.ps1." }
            }
             if ($moduleFile -eq "TMS_Settings.ps1") {
                 if (-not (Get-Command Update-TariffMargin -ErrorAction SilentlyContinue)) { throw "Required function 'Update-TariffMargin' not found immediately after loading TMS_Settings.ps1." }
             }
             # <<< MODIFICATION: Updated checks for split helper files >>>
             if ($moduleFile -eq "TMS_Helpers_General.ps1") {
                  if (-not (Get-Command Get-PermittedKeys -ErrorAction SilentlyContinue)) { throw "Required function 'Get-PermittedKeys' not found in TMS_Helpers_General.ps1." }
                  if (-not (Get-Command Select-CsvFile -ErrorAction SilentlyContinue)) { throw "Required function 'Select-CsvFile' not found in TMS_Helpers_General.ps1." }
                  if (-not (Get-Command Open-FileExplorer -ErrorAction SilentlyContinue)) { throw "Required function 'Open-FileExplorer' not found in TMS_Helpers_General.ps1." }
                  if (-not (Get-Command Load-KeysFromFolder -ErrorAction SilentlyContinue)) { throw "Required function 'Load-KeysFromFolder' not found in TMS_Helpers_General.ps1." }
             }
             if ($moduleFile -eq "TMS_GUI_Helpers.ps1") {
                  if (-not (Get-Command Populate-TariffListBox -ErrorAction SilentlyContinue)) { throw "Required function 'Populate-TariffListBox' not found in TMS_GUI_Helpers.ps1." }
                  if (-not (Get-Command Populate-ReportTariffListBoxes -ErrorAction SilentlyContinue)) { throw "Required function 'Populate-ReportTariffListBoxes' not found in TMS_GUI_Helpers.ps1." }
             }
             if ($moduleFile -eq "TMS_Helpers_Central.ps1") {
                 if (-not (Get-Command Invoke-CentralTransportApi -ErrorAction SilentlyContinue)) { throw "Required function 'Invoke-CentralTransportApi' not found in TMS_Helpers_Central.ps1."}
                 if (-not (Get-Command Load-And-Normalize-CentralData -ErrorAction SilentlyContinue)) { throw "Required function 'Load-And-Normalize-CentralData' not found in TMS_Helpers_Central.ps1."}
             }
             if ($moduleFile -eq "TMS_Helpers_SAIA.ps1") {
                 if (-not (Get-Command Invoke-SAIAApi -ErrorAction SilentlyContinue)) { throw "Required function 'Invoke-SAIAApi' not found in TMS_Helpers_SAIA.ps1."}
                 if (-not (Get-Command Load-And-Normalize-SAIAData -ErrorAction SilentlyContinue)) { throw "Required function 'Load-And-Normalize-SAIAData' not found in TMS_Helpers_SAIA.ps1."}
             }
             if ($moduleFile -eq "TMS_Helpers_RL.ps1") {
                 if (-not (Get-Command Invoke-RLApi -ErrorAction SilentlyContinue)) { throw "Required function 'Invoke-RLApi' not found in TMS_Helpers_RL.ps1."}
                 if (-not (Get-Command Load-And-Normalize-RLData -ErrorAction SilentlyContinue)) { throw "Required function 'Load-And-Normalize-RLData' not found in TMS_Helpers_RL.ps1."}
             }
             # <<< END MODIFICATION >>>
             if ($moduleFile -eq "TMS_Carrier_Central.ps1") { if (-not (Get-Command Run-CentralComparisonReportGUI -ErrorAction SilentlyContinue)) { throw "Required function 'Run-CentralComparisonReportGUI' not found." } }
             if ($moduleFile -eq "TMS_Carrier_SAIA.ps1") { if (-not (Get-Command Run-SAIAComparisonReportGUI -ErrorAction SilentlyContinue)) { throw "Required function 'Run-SAIAComparisonReportGUI' not found." } }
             if ($moduleFile -eq "TMS_Carrier_RL.ps1") { if (-not (Get-Command Run-RLComparisonReportGUI -ErrorAction SilentlyContinue)) { throw "Required function 'Run-RLComparisonReportGUI' not found." } }
             if ($moduleFile -eq "TMS_Reports.ps1") {
                 if (-not (Get-Command Run-CrossCarrierASPAnalysisGUI -ErrorAction SilentlyContinue)) { throw "Required function 'Run-CrossCarrierASPAnalysisGUI' not found." }
                 if (-not (Get-Command Run-MarginsByHistoryAnalysisGUI -ErrorAction SilentlyContinue)) { throw "Required function 'Run-MarginsByHistoryAnalysisGUI' not found." } 
             }
        } catch { Write-Error "ERROR loading module '$moduleFile': $($_.Exception.Message)"; throw $_ }
    }
    # --- Final Verification AFTER loop (Add more as needed) ---
    if (-not (Get-Command Update-TariffMargin -ErrorAction SilentlyContinue)) { throw "Required function 'Update-TariffMargin' not found AFTER loop." }
    if (-not (Get-Command Populate-TariffListBox -ErrorAction SilentlyContinue)) { throw "Required function 'Populate-TariffListBox' from TMS_GUI_Helpers.ps1 not found AFTER loop."}
    if (-not (Get-Command Load-KeysFromFolder -ErrorAction SilentlyContinue)) { throw "Required function 'Load-KeysFromFolder' from TMS_Helpers_General.ps1 not found AFTER loop."}
    if (-not (Get-Command Invoke-CentralTransportApi -ErrorAction SilentlyContinue)) { throw "Required function 'Invoke-CentralTransportApi' from TMS_Helpers_Central.ps1 not found AFTER loop."}


    Write-Verbose "Module loading complete."
} catch {
     $errorMessage = "FATAL: Failed to load a module or required function. GUI cannot start.`nError: $($_.Exception.Message)"; Write-Error $errorMessage; [System.Windows.Forms.MessageBox]::Show($errorMessage, "Module Load Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); Exit 1
}

# --- Resolve Full Paths for Data Folders ---
$CentralKeysFolderName = $script:defaultCentralKeysFolderName; $SAIAKeysFolderName = $script:defaultSAIAKeysFolderName; $RLKeysFolderName = $script:defaultRLKeysFolderName
$UserAccountsFolderName = $script:defaultUserAccountsFolderName; $CustomerAccountsFolderName = $script:defaultCustomerAccountsFolderName
$ReportsBaseFolderName = $script:defaultReportsBaseFolderName; $ShipmentDataFolderName = $script:defaultShipmentDataFolderName
$script:centralKeysFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $CentralKeysFolderName; $script:saiaKeysFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $SAIAKeysFolderName; $script:rlKeysFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $RLKeysFolderName
$script:userAccountsFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $UserAccountsFolderName; $script:customerAccountsFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $CustomerAccountsFolderName
$script:reportsBaseFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $ReportsBaseFolderName; $script:shipmentDataFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $ShipmentDataFolderName

# --- Ensure Required Base Folders Exist ---
Write-Verbose "Ensuring base data directories exist..."
try { Ensure-DirectoryExists -Path $script:centralKeysFolderPath; Ensure-DirectoryExists -Path $script:saiaKeysFolderPath; Ensure-DirectoryExists -Path $script:rlKeysFolderPath; Ensure-DirectoryExists -Path $script:userAccountsFolderPath; Ensure-DirectoryExists -Path $script:customerAccountsFolderPath; Ensure-DirectoryExists -Path $script:reportsBaseFolderPath; Ensure-DirectoryExists -Path $script:shipmentDataFolderPath }
catch { $errorMessage = "FATAL: Failed to ensure required data directories exist.`nError: $($_.Exception.Message)"; Write-Error $errorMessage; [System.Windows.Forms.MessageBox]::Show($errorMessage, "Directory Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); Exit 1 }
Write-Verbose "Base directory check complete."

# --- Pre-load All Carrier Keys/Data ---
Write-Verbose "Loading all available carrier keys/accounts/margins..."
$script:allCentralKeys = @{}; $script:allSAIAKeys = @{}; $script:allRLKeys = @{}
try {
    $script:allCentralKeys = Load-KeysFromFolder -KeysFolderPath $script:centralKeysFolderPath -CarrierName "Central Transport"
    Write-Host "DEBUG GUI (Startup): Loaded Central Keys. Type: $($script:allCentralKeys.GetType().FullName). Count: $($script:allCentralKeys.Count)."
    if ($script:allCentralKeys -isnot [hashtable]) { Write-Warning "DEBUG GUI (Startup): allCentralKeys is NOT a hashtable after Load-KeysFromFolder!" }
    $script:allSAIAKeys = Load-KeysFromFolder -KeysFolderPath $script:saiaKeysFolderPath -CarrierName "SAIA"
    Write-Host "DEBUG GUI (Startup): Loaded SAIA Keys. Type: $($script:allSAIAKeys.GetType().FullName). Count: $($script:allSAIAKeys.Count)."
    if ($script:allSAIAKeys -isnot [hashtable]) { Write-Warning "DEBUG GUI (Startup): allSAIAKeys is NOT a hashtable after Load-KeysFromFolder!" }
    $script:allRLKeys = Load-KeysFromFolder -KeysFolderPath $script:rlKeysFolderPath -CarrierName "RL Carriers"
    Write-Host "DEBUG GUI (Startup): Loaded R+L Keys. Type: $($script:allRLKeys.GetType().FullName). Count: $($script:allRLKeys.Count)."
    if ($script:allRLKeys -isnot [hashtable]) { Write-Warning "DEBUG GUI (Startup): allRLKeys is NOT a hashtable after Load-KeysFromFolder!" }
} catch {
    $loadErrorMsg = "FATAL ERROR during Load-KeysFromFolder: $($_.Exception.Message). Check function availability and key file paths."
    Write-Error $loadErrorMsg
    [System.Windows.Forms.MessageBox]::Show($loadErrorMsg, "Key Loading Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
}
Write-Verbose "Key/Account/Margin loading complete."

# --- Pre-load All Customer Profiles ---
Write-Verbose "Loading all available customer profiles..."
$script:allCustomerProfiles = Load-AllCustomerProfiles -UserAccountsFolderPath $script:customerAccountsFolderPath
Write-Verbose "Customer profile loading complete."
if($script:allCustomerProfiles.Count -eq 0){ Write-Warning "DEBUG GUI (Startup): No customer profiles loaded!"}
else { Write-Host "DEBUG GUI (Startup): $($script:allCustomerProfiles.Count) customer profiles loaded." }


$script:currentUserProfile = $null; $script:selectedCustomerProfile = $null; $script:currentUserReportsFolder = $null
$fontRegular = New-Object System.Drawing.Font("Segoe UI", 9); $fontBold = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $fontTitle = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold); $fontMono = New-Object System.Drawing.Font("Consolas", 9)
$colorBackground = [System.Drawing.Color]::FromArgb(240, 242, 245); $colorPanel = [System.Drawing.Color]::White; $colorPrimary = [System.Drawing.Color]::FromArgb(0, 120, 215); $colorPrimaryLight = [System.Drawing.Color]::FromArgb(100, 180, 240)
$colorText = [System.Drawing.Color]::FromArgb(30, 30, 30); $colorTextLight = [System.Drawing.Color]::FromArgb(100, 100, 100); $colorButtonBack = $colorPrimary; $colorButtonFore = [System.Drawing.Color]::White; $colorButtonBorder = [System.Drawing.Color]::FromArgb(0, 90, 180)
$colorInputBorder = [System.Drawing.Color]::FromArgb(200, 200, 200); $paddingSmall = New-Object System.Windows.Forms.Padding(5); $paddingMedium = New-Object System.Windows.Forms.Padding(10)

$mainForm = New-Object System.Windows.Forms.Form; $mainForm.Text = "TMS GUI Tool (Broker Mode)"; $mainForm.Size = New-Object System.Drawing.Size(850, 700); $mainForm.MinimumSize = New-Object System.Drawing.Size(800, 600); $mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen; $mainForm.MaximizeBox = $true; $mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable; $mainForm.BackColor = $colorBackground; $mainForm.Font = $fontRegular

$loginPanel = New-Object System.Windows.Forms.Panel; $loginPanel.Location = New-Object System.Drawing.Point(10, 10); $loginPanel.Size = New-Object System.Drawing.Size(320, 170); $loginPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle; $loginPanel.Anchor = [System.Windows.Forms.AnchorStyles]::None; $loginPanel.BackColor = $colorPanel; $loginPanel.Padding = $paddingMedium
$loginTitleLabel = New-Object System.Windows.Forms.Label; $loginTitleLabel.Text = "Broker Login"; $loginTitleLabel.Font = $fontTitle; $loginTitleLabel.ForeColor = $colorPrimary; $loginTitleLabel.AutoSize = $true; $loginTitleLabel.Location = New-Object System.Drawing.Point(10, 10)
$labelUsername = New-Object System.Windows.Forms.Label; $labelUsername.Text = "Username:"; $labelUsername.Location = New-Object System.Drawing.Point(10, 55); $labelUsername.AutoSize = $true; $labelUsername.ForeColor = $colorText
$textboxUsername = New-Object System.Windows.Forms.TextBox; $textboxUsername.Location = New-Object System.Drawing.Point(95, 52); $textboxUsername.Size = New-Object System.Drawing.Size(200, 23); $textboxUsername.Font = $fontRegular; $textboxUsername.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelPassword = New-Object System.Windows.Forms.Label; $labelPassword.Text = "Password:"; $labelPassword.Location = New-Object System.Drawing.Point(10, 88); $labelPassword.AutoSize = $true; $labelPassword.ForeColor = $colorText
$textboxPassword = New-Object System.Windows.Forms.TextBox; $textboxPassword.Location = New-Object System.Drawing.Point(95, 85); $textboxPassword.Size = New-Object System.Drawing.Size(200, 23); $textboxPassword.UseSystemPasswordChar = $true; $textboxPassword.Font = $fontRegular; $textboxPassword.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$buttonLogin = New-Object System.Windows.Forms.Button; $buttonLogin.Text = "Login"; $buttonLogin.Location = New-Object System.Drawing.Point(120, 125); $buttonLogin.Size = New-Object System.Drawing.Size(80, 30); $buttonLogin.Font = $fontBold; $buttonLogin.BackColor = $colorButtonBack; $buttonLogin.ForeColor = $colorButtonFore; $buttonLogin.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $buttonLogin.FlatAppearance.BorderSize = 1; $buttonLogin.FlatAppearance.BorderColor = $colorButtonBorder
$loginPanel.Controls.AddRange(@($loginTitleLabel, $labelUsername, $textboxUsername, $labelPassword, $textboxPassword, $buttonLogin)); $mainForm.Controls.Add($loginPanel)
$mainForm.Add_Resize({ if ($loginPanel.Visible) { $loginPanel.Left = ($mainForm.ClientSize.Width - $loginPanel.Width) / 2; $loginPanel.Top = ($mainForm.ClientSize.Height - $loginPanel.Height) / 3 } })

$statusBar = New-Object System.Windows.Forms.StatusBar; $statusBar.Text = "Ready. Please login."; $statusBar.Font = $fontRegular
$tabControlMain = New-Object System.Windows.Forms.TabControl; $tabControlMain.Location = New-Object System.Drawing.Point(10, 10); $tabControlMain.Size = New-Object System.Drawing.Size(($mainForm.ClientSize.Width - 20), ($mainForm.ClientSize.Height - 50)); $tabControlMain.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right; $tabControlMain.Visible = $false; $tabControlMain.Font = $fontRegular; $tabControlMain.Padding = New-Object System.Drawing.Point(10, 5); $mainForm.Controls.Add($tabControlMain)

$tabPageQuote = New-Object System.Windows.Forms.TabPage; $tabPageQuote.Text = "Quote"; $tabPageQuote.BackColor = $colorPanel; $tabPageQuote.Padding = $paddingMedium; $tabControlMain.Controls.Add($tabPageQuote)
$tabPageSettings = New-Object System.Windows.Forms.TabPage; $tabPageSettings.Text = "Settings"; $tabPageSettings.BackColor = $colorPanel; $tabPageSettings.Padding = $paddingMedium; $tabControlMain.Controls.Add($tabPageSettings)
$tabPageReports = New-Object System.Windows.Forms.TabPage; $tabPageReports.Text = "Reports"; $tabPageReports.BackColor = $colorPanel; $tabPageReports.Padding = $paddingMedium; $tabControlMain.Controls.Add($tabPageReports) 

# ============================================================
# Single Quote UI Section (Inside $tabPageQuote) 
# ============================================================
$singleQuotePanel = New-Object System.Windows.Forms.Panel; $singleQuotePanel.Dock = [System.Windows.Forms.DockStyle]::Fill; $singleQuotePanel.BackColor = $colorPanel
$labelSelectCustomer_Quote = New-Object System.Windows.Forms.Label; $labelSelectCustomer_Quote.Text = "Select Customer:"; $labelSelectCustomer_Quote.Location = New-Object System.Drawing.Point(550, 15); $labelSelectCustomer_Quote.AutoSize = $true; $labelSelectCustomer_Quote.Font = $fontBold; $labelSelectCustomer_Quote.ForeColor = $colorText
$comboBoxSelectCustomer_Quote = New-Object System.Windows.Forms.ComboBox; $comboBoxSelectCustomer_Quote.Location = New-Object System.Drawing.Point(550, 35); $comboBoxSelectCustomer_Quote.Size = New-Object System.Drawing.Size(200, 23); $comboBoxSelectCustomer_Quote.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; $comboBoxSelectCustomer_Quote.Enabled = $false; $comboBoxSelectCustomer_Quote.Font = $fontRegular
[int]$col1X = 15; [int]$col2X = 300; [int]$labelWidth = 90; [int]$textBoxWidth = 160; [int]$rowHeight = 30
$labelOriginHeader = New-Object System.Windows.Forms.Label; $labelOriginHeader.Text = "Origin Details:"; $labelOriginHeader.Location = New-Object System.Drawing.Point($col1X, 15); $labelOriginHeader.Font = $fontBold; $labelOriginHeader.AutoSize = $true; $labelOriginHeader.ForeColor = $colorPrimary
$labelDestHeader = New-Object System.Windows.Forms.Label; $labelDestHeader.Text = "Destination Details:"; $labelDestHeader.Location = New-Object System.Drawing.Point($col2X, 15); $labelDestHeader.Font = $fontBold; $labelDestHeader.AutoSize = $true; $labelDestHeader.ForeColor = $colorPrimary
$labelOriginZip = New-Object System.Windows.Forms.Label; $labelOriginZip.Text = "ZIP Code:"; $labelOriginZip.Location = New-Object System.Drawing.Point($col1X, (15 + $rowHeight)); $labelOriginZip.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelOriginZip.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $labelOriginZip.ForeColor = $colorText
$textboxOriginZip = New-Object System.Windows.Forms.TextBox; $textboxOriginZip.Location = New-Object System.Drawing.Point(($col1X + $labelWidth + 5), (15 + $rowHeight)); $textboxOriginZip.Size = New-Object System.Drawing.Size($textBoxWidth, 23); $textboxOriginZip.MaxLength = 5; $textboxOriginZip.Font = $fontRegular; $textboxOriginZip.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelDestZip = New-Object System.Windows.Forms.Label; $labelDestZip.Text = "ZIP Code:"; $labelDestZip.Location = New-Object System.Drawing.Point($col2X, (15 + $rowHeight)); $labelDestZip.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelDestZip.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $labelDestZip.ForeColor = $colorText
$textboxDestZip = New-Object System.Windows.Forms.TextBox; $textboxDestZip.Location = New-Object System.Drawing.Point(($col2X + $labelWidth + 5), (15 + $rowHeight)); $textboxDestZip.Size = New-Object System.Drawing.Size($textBoxWidth, 23); $textboxDestZip.MaxLength = 5; $textboxDestZip.Font = $fontRegular; $textboxDestZip.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelOriginCity = New-Object System.Windows.Forms.Label; $labelOriginCity.Text = "City:"; $labelOriginCity.Location = New-Object System.Drawing.Point($col1X, (15 + (2 * $rowHeight))); $labelOriginCity.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelOriginCity.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $labelOriginCity.ForeColor = $colorText
$textboxOriginCity = New-Object System.Windows.Forms.TextBox; $textboxOriginCity.Location = New-Object System.Drawing.Point(($col1X + $labelWidth + 5), (15 + (2 * $rowHeight))); $textboxOriginCity.Size = New-Object System.Drawing.Size($textBoxWidth, 23); $textboxOriginCity.Font = $fontRegular; $textboxOriginCity.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelDestCity = New-Object System.Windows.Forms.Label; $labelDestCity.Text = "City:"; $labelDestCity.Location = New-Object System.Drawing.Point($col2X, (15 + (2 * $rowHeight))); $labelDestCity.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelDestCity.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $labelDestCity.ForeColor = $colorText
$textboxDestCity = New-Object System.Windows.Forms.TextBox; $textboxDestCity.Location = New-Object System.Drawing.Point(($col2X + $labelWidth + 5), (15 + (2 * $rowHeight))); $textboxDestCity.Size = New-Object System.Drawing.Size($textBoxWidth, 23); $textboxDestCity.Font = $fontRegular; $textboxDestCity.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelOriginState = New-Object System.Windows.Forms.Label; $labelOriginState.Text = "State (2 Ltr):"; $labelOriginState.Location = New-Object System.Drawing.Point($col1X, (15 + (3 * $rowHeight))); $labelOriginState.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelOriginState.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $labelOriginState.ForeColor = $colorText
$textboxOriginState = New-Object System.Windows.Forms.TextBox; $textboxOriginState.Location = New-Object System.Drawing.Point(($col1X + $labelWidth + 5), (15 + (3 * $rowHeight))); $textboxOriginState.Size = New-Object System.Drawing.Size(50, 23); $textboxOriginState.MaxLength = 2; $textboxOriginState.CharacterCasing = [System.Windows.Forms.CharacterCasing]::Upper; $textboxOriginState.Font = $fontRegular; $textboxOriginState.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelDestState = New-Object System.Windows.Forms.Label; $labelDestState.Text = "State (2 Ltr):"; $labelDestState.Location = New-Object System.Drawing.Point($col2X, (15 + (3 * $rowHeight))); $labelDestState.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelDestState.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $labelDestState.ForeColor = $colorText
$textboxDestState = New-Object System.Windows.Forms.TextBox; $textboxDestState.Location = New-Object System.Drawing.Point(($col2X + $labelWidth + 5), (15 + (3 * $rowHeight))); $textboxDestState.Size = New-Object System.Drawing.Size(50, 23); $textboxDestState.MaxLength = 2; $textboxDestState.CharacterCasing = [System.Windows.Forms.CharacterCasing]::Upper; $textboxDestState.Font = $fontRegular; $textboxDestState.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelWeight = New-Object System.Windows.Forms.Label; $labelWeight.Text = "Weight (lbs):"; $labelWeight.Location = New-Object System.Drawing.Point($col1X, (15 + (4 * $rowHeight))); $labelWeight.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelWeight.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $labelWeight.ForeColor = $colorText
$textboxWeight = New-Object System.Windows.Forms.TextBox; $textboxWeight.Location = New-Object System.Drawing.Point(($col1X + $labelWidth + 5), (15 + (4 * $rowHeight))); $textboxWeight.Size = New-Object System.Drawing.Size($textBoxWidth, 23); $textboxWeight.Font = $fontRegular; $textboxWeight.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelClass = New-Object System.Windows.Forms.Label; $labelClass.Text = "Class:"; $labelClass.Location = New-Object System.Drawing.Point($col2X, (15 + (4 * $rowHeight))); $labelClass.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelClass.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $labelClass.ForeColor = $colorText
$textboxClass = New-Object System.Windows.Forms.TextBox; $textboxClass.Location = New-Object System.Drawing.Point(($col2X + $labelWidth + 5), (15 + (4 * $rowHeight))); $textboxClass.Size = New-Object System.Drawing.Size($textBoxWidth, 23); $textboxClass.Font = $fontRegular; $textboxClass.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$groupBoxOptional = New-Object System.Windows.Forms.GroupBox; $groupBoxOptional.Text = "Optional Details"; $groupBoxOptional.Location = New-Object System.Drawing.Point($col1X, (15 + (5 * $rowHeight) + 10)); $groupBoxOptional.Size = New-Object System.Drawing.Size(510, 60); $groupBoxOptional.ForeColor = $colorText; $groupBoxOptional.Font = $fontRegular
$labelLength = New-Object System.Windows.Forms.Label; $labelLength.Text = "L:"; $labelLength.Location = New-Object System.Drawing.Point(15, 25); $labelLength.AutoSize=$true
$textboxLength = New-Object System.Windows.Forms.TextBox; $textboxLength.Location = New-Object System.Drawing.Point(35, 22); $textboxLength.Size = New-Object System.Drawing.Size(45, 23); $textboxLength.Text = "1.0"; $textboxLength.Font = $fontRegular; $textboxLength.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelWidthOpt = New-Object System.Windows.Forms.Label; $labelWidthOpt.Text = "W:"; $labelWidthOpt.Location = New-Object System.Drawing.Point(90, 25); $labelWidthOpt.AutoSize=$true
$textboxWidthOpt = New-Object System.Windows.Forms.TextBox; $textboxWidthOpt.Location = New-Object System.Drawing.Point(115, 22); $textboxWidthOpt.Size = New-Object System.Drawing.Size(45, 23); $textboxWidthOpt.Text = "1.0"; $textboxWidthOpt.Font = $fontRegular; $textboxWidthOpt.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelHeight = New-Object System.Windows.Forms.Label; $labelHeight.Text = "H:"; $labelHeight.Location = New-Object System.Drawing.Point(170, 25); $labelHeight.AutoSize=$true
$textboxHeight = New-Object System.Windows.Forms.TextBox; $textboxHeight.Location = New-Object System.Drawing.Point(195, 22); $textboxHeight.Size = New-Object System.Drawing.Size(45, 23); $textboxHeight.Text = "1.0"; $textboxHeight.Font = $fontRegular; $textboxHeight.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelDimsUnit = New-Object System.Windows.Forms.Label; $labelDimsUnit.Text = "(inches)"; $labelDimsUnit.Location = New-Object System.Drawing.Point(250, 25); $labelDimsUnit.AutoSize=$true; $labelDimsUnit.ForeColor = $colorTextLight
$labelDeclaredValue = New-Object System.Windows.Forms.Label; $labelDeclaredValue.Text = "Declared Val ($):"; $labelDeclaredValue.Location = New-Object System.Drawing.Point(310, 25); $labelDeclaredValue.AutoSize=$true
$textboxDeclaredValue = New-Object System.Windows.Forms.TextBox; $textboxDeclaredValue.Location = New-Object System.Drawing.Point(420, 22); $textboxDeclaredValue.Size = New-Object System.Drawing.Size(80, 23); $textboxDeclaredValue.Text = "0.00"; $textboxDeclaredValue.Font = $fontRegular; $textboxDeclaredValue.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$groupBoxOptional.Controls.AddRange(@($labelLength, $textboxLength, $labelWidthOpt, $textboxWidthOpt, $labelHeight, $textboxHeight, $labelDimsUnit, $labelDeclaredValue, $textboxDeclaredValue))
$buttonGetQuote = New-Object System.Windows.Forms.Button; $buttonGetQuote.Text = "Get Quote"; $buttonGetQuote.Location = New-Object System.Drawing.Point(550, 185); $buttonGetQuote.Size = New-Object System.Drawing.Size(110, 35); $buttonGetQuote.Font = $fontBold; $buttonGetQuote.BackColor = $colorButtonBack; $buttonGetQuote.ForeColor = $colorButtonFore; $buttonGetQuote.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $buttonGetQuote.FlatAppearance.BorderSize = 1; $buttonGetQuote.FlatAppearance.BorderColor = $colorButtonBorder
$labelResults = New-Object System.Windows.Forms.Label; $labelResults.Text = "Quote Results:"; $labelResults.Location = New-Object System.Drawing.Point($col1X, (15 + (5 * $rowHeight) + 85)); $labelResults.Font = $fontBold; $labelResults.AutoSize = $true; $labelResults.ForeColor = $colorPrimary
$textboxResults = New-Object System.Windows.Forms.TextBox; $textboxResults.Location = New-Object System.Drawing.Point($col1X, (15 + (5 * $rowHeight) + 105)); $textboxResults.Size = New-Object System.Drawing.Size(($singleQuotePanel.ClientSize.Width - $col1X - 15), 140); $textboxResults.Multiline = $true; $textboxResults.ReadOnly = $true; $textboxResults.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical; $textboxResults.Font = $fontMono; $textboxResults.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right; $textboxResults.BackColor = $colorPanel; $textboxResults.ForeColor = $colorText
$singleQuotePanel.Controls.AddRange(@( $labelSelectCustomer_Quote, $comboBoxSelectCustomer_Quote, $labelOriginHeader, $labelDestHeader, $labelOriginZip, $textboxOriginZip, $labelDestZip, $textboxDestZip, $labelOriginCity, $textboxOriginCity, $labelDestCity, $textboxDestCity, $labelOriginState, $textboxOriginState, $labelDestState, $textboxDestState, $labelWeight, $textboxWeight, $labelClass, $textboxClass, $groupBoxOptional, $buttonGetQuote, $labelResults, $textboxResults ))
$tabPageQuote.Controls.Add($singleQuotePanel)

# ============================================================
# Settings UI Section (Inside $tabPageSettings) 
# ============================================================
$settingsPanel = New-Object System.Windows.Forms.Panel; $settingsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill; $settingsPanel.BackColor = $colorPanel
$tabPageSettings.Controls.Add($settingsPanel)
$labelSelectCustomer_Settings = New-Object System.Windows.Forms.Label; $labelSelectCustomer_Settings.Text = "Select Customer:"; $labelSelectCustomer_Settings.Location = New-Object System.Drawing.Point(330, 15); $labelSelectCustomer_Settings.AutoSize = $true; $labelSelectCustomer_Settings.Font = $fontBold; $labelSelectCustomer_Settings.ForeColor = $colorText
$comboBoxSelectCustomer_Settings = New-Object System.Windows.Forms.ComboBox; $comboBoxSelectCustomer_Settings.Location = New-Object System.Drawing.Point(330, 35); $comboBoxSelectCustomer_Settings.Size = New-Object System.Drawing.Size(250, 23); $comboBoxSelectCustomer_Settings.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; $comboBoxSelectCustomer_Settings.Enabled = $false; $comboBoxSelectCustomer_Settings.Font = $fontRegular
$groupBoxCarrierSelect = New-Object System.Windows.Forms.GroupBox; $groupBoxCarrierSelect.Text = "Select Carrier"; $groupBoxCarrierSelect.Location = New-Object System.Drawing.Point(15, 15); $groupBoxCarrierSelect.Size = New-Object System.Drawing.Size(300, 55); $groupBoxCarrierSelect.Font = $fontRegular; $groupBoxCarrierSelect.ForeColor = $colorText
$radioCentral = New-Object System.Windows.Forms.RadioButton; $radioCentral.Text = "Central"; $radioCentral.Location = New-Object System.Drawing.Point(15, 22); $radioCentral.AutoSize = $true; $radioCentral.Checked = $true; $radioCentral.Font = $fontRegular; $radioCentral.ForeColor = $colorText
$radioSAIA = New-Object System.Windows.Forms.RadioButton; $radioSAIA.Text = "SAIA"; $radioSAIA.Location = New-Object System.Drawing.Point(100, 22); $radioSAIA.AutoSize = $true; $radioSAIA.Font = $fontRegular; $radioSAIA.ForeColor = $colorText
$radioRL = New-Object System.Windows.Forms.RadioButton; $radioRL.Text = "R+L"; $radioRL.Location = New-Object System.Drawing.Point(180, 22); $radioRL.AutoSize = $true; $radioRL.Font = $fontRegular; $radioRL.ForeColor = $colorText
$groupBoxCarrierSelect.Controls.AddRange(@($radioCentral, $radioSAIA, $radioRL))
$labelTariffList = New-Object System.Windows.Forms.Label; $labelTariffList.Text = "Permitted Tariffs && Margins (for Selected Customer):"; $labelTariffList.Location = New-Object System.Drawing.Point(15, 80); $labelTariffList.AutoSize = $true; $labelTariffList.Font = $fontBold; $labelTariffList.ForeColor = $colorPrimary
$listBoxTariffs = New-Object System.Windows.Forms.ListBox; $listBoxTariffs.Location = New-Object System.Drawing.Point(15, 105); $listBoxTariffs.Size = New-Object System.Drawing.Size(300, 220); $listBoxTariffs.Font = $fontMono; $listBoxTariffs.IntegralHeight = $false; $listBoxTariffs.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$groupBoxSetMargin = New-Object System.Windows.Forms.GroupBox; $groupBoxSetMargin.Text = "Set Margin for Selected Tariff"; $groupBoxSetMargin.Location = New-Object System.Drawing.Point(330, 75); $groupBoxSetMargin.Size = New-Object System.Drawing.Size(250, 130); $groupBoxSetMargin.Font = $fontRegular; $groupBoxSetMargin.ForeColor = $colorText
$labelSelectedTariff = New-Object System.Windows.Forms.Label; $labelSelectedTariff.Text = "Selected: (None)"; $labelSelectedTariff.Location = New-Object System.Drawing.Point(15, 28); $labelSelectedTariff.AutoSize = $true; $labelSelectedTariff.Font = $fontBold; $labelSelectedTariff.ForeColor = $colorText
$labelNewMargin = New-Object System.Windows.Forms.Label; $labelNewMargin.Text = "New Margin %:"; $labelNewMargin.Location = New-Object System.Drawing.Point(15, 60); $labelNewMargin.AutoSize = $true; $labelNewMargin.ForeColor = $colorText
$textBoxNewMargin = New-Object System.Windows.Forms.TextBox; $textBoxNewMargin.Location = New-Object System.Drawing.Point(120, 57); $textBoxNewMargin.Size = New-Object System.Drawing.Size(70, 23); $textBoxNewMargin.Enabled = $false; $textBoxNewMargin.Font = $fontRegular; $textBoxNewMargin.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$buttonSetMargin = New-Object System.Windows.Forms.Button; $buttonSetMargin.Text = "Set Margin"; $buttonSetMargin.Location = New-Object System.Drawing.Point(75, 90); $buttonSetMargin.Size = New-Object System.Drawing.Size(100, 30); $buttonSetMargin.Enabled = $false; $buttonSetMargin.Font = $fontBold; $buttonSetMargin.BackColor = $colorButtonBack; $buttonSetMargin.ForeColor = $colorButtonFore; $buttonSetMargin.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $buttonSetMargin.FlatAppearance.BorderSize = 1; $buttonSetMargin.FlatAppearance.BorderColor = $colorButtonBorder
$groupBoxSetMargin.Controls.AddRange(@($labelSelectedTariff, $labelNewMargin, $textBoxNewMargin, $buttonSetMargin))
$settingsPanel.Controls.AddRange(@( $labelSelectCustomer_Settings, $comboBoxSelectCustomer_Settings, $groupBoxCarrierSelect, $labelTariffList, $listBoxTariffs, $groupBoxSetMargin ))

# ============================================================
# Reports UI Section (Inside $tabPageReports) 
# ============================================================
$reportsPanel = New-Object System.Windows.Forms.Panel; $reportsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill; $reportsPanel.BackColor = $colorPanel
$tabPageReports.Controls.Add($reportsPanel)

[int]$reportsPanelX = 15; [int]$reportsPanelY = 15
[int]$reportsControlSpacing = 35; [int]$reportsLabelWidth = 150; [int]$reportsInputWidth = 250 
[int]$reportsListBoxWidth = 250; [int]$reportsListBoxHeight = 100
[int]$groupBoxAspInputHeight = 55; [int]$checkBoxApplyMarginsHeight = 25 
[int]$verticalPaddingBetweenControls = 10 

$currentY = $reportsPanelY 

$labelSelectCustomer_Reports = New-Object System.Windows.Forms.Label; $labelSelectCustomer_Reports.Text = "Select Customer:"; $labelSelectCustomer_Reports.Location = New-Object System.Drawing.Point($reportsPanelX, $currentY); $labelSelectCustomer_Reports.AutoSize = $true; $labelSelectCustomer_Reports.Font = $fontBold; $labelSelectCustomer_Reports.ForeColor = $colorText
$comboBoxSelectCustomer_Reports = New-Object System.Windows.Forms.ComboBox; $comboBoxSelectCustomer_Reports.Location = New-Object System.Drawing.Point(($reportsPanelX + $reportsLabelWidth + 5), ($currentY - 3)); $comboBoxSelectCustomer_Reports.Size = New-Object System.Drawing.Size($reportsInputWidth, 23); $comboBoxSelectCustomer_Reports.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; $comboBoxSelectCustomer_Reports.Enabled = $false; $comboBoxSelectCustomer_Reports.Font = $fontRegular
$currentY += $reportsControlSpacing

$labelReportType = New-Object System.Windows.Forms.Label; $labelReportType.Text = "Select Report Type:"; $labelReportType.Location = New-Object System.Drawing.Point($reportsPanelX, $currentY); $labelReportType.AutoSize = $true; $labelReportType.Font = $fontBold; $labelReportType.ForeColor = $colorText
$comboBoxReportType = New-Object System.Windows.Forms.ComboBox; $comboBoxReportType.Location = New-Object System.Drawing.Point(($reportsPanelX + $reportsLabelWidth + 5), ($currentY - 3)); $comboBoxReportType.Size = New-Object System.Drawing.Size($reportsInputWidth, 23); $comboBoxReportType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; $comboBoxReportType.Enabled = $false; $comboBoxReportType.Font = $fontRegular
$comboBoxReportType.Items.AddRange(@( "Carrier Comparison", "Avg Required Margin", "Required Margin for ASP", "Cross-Carrier ASP Analysis", "Margins by History" ))
$currentY += $reportsControlSpacing # Base Y for next set of controls

# Define controls that might shift
$groupBoxReportCarrierSelect = New-Object System.Windows.Forms.GroupBox; $groupBoxReportCarrierSelect.Text = "Select Carrier"; $groupBoxReportCarrierSelect.Size = New-Object System.Drawing.Size(300, 55); $groupBoxReportCarrierSelect.Font = $fontRegular; $groupBoxReportCarrierSelect.ForeColor = $colorText; $groupBoxReportCarrierSelect.Visible = $false 
$radioReportCentral = New-Object System.Windows.Forms.RadioButton; $radioReportCentral.Text = "Central"; $radioReportCentral.Location = New-Object System.Drawing.Point(15, 22); $radioReportCentral.AutoSize = $true; $radioReportCentral.Checked = $true; $radioReportCentral.Font = $fontRegular; $radioReportCentral.ForeColor = $colorText
$radioReportSAIA = New-Object System.Windows.Forms.RadioButton; $radioReportSAIA.Text = "SAIA"; $radioReportSAIA.Location = New-Object System.Drawing.Point(100, 22); $radioReportSAIA.AutoSize = $true; $radioReportSAIA.Font = $fontRegular; $radioReportSAIA.ForeColor = $colorText
$radioReportRL = New-Object System.Windows.Forms.RadioButton; $radioReportRL.Text = "R+L"; $radioReportRL.Location = New-Object System.Drawing.Point(180, 22); $radioReportRL.AutoSize = $true; $radioReportRL.Font = $fontRegular; $radioReportRL.ForeColor = $colorText
$groupBoxReportCarrierSelect.Controls.AddRange(@($radioReportCentral, $radioReportSAIA, $radioReportRL))

$groupBoxReportTariffSelect = New-Object System.Windows.Forms.GroupBox; $groupBoxReportTariffSelect.Text = "Select Tariff(s)"; $groupBoxReportTariffSelect.Size = New-Object System.Drawing.Size(550, 140); $groupBoxReportTariffSelect.Font = $fontRegular; $groupBoxReportTariffSelect.ForeColor = $colorText; $groupBoxReportTariffSelect.Visible = $false 
$labelReportTariff1 = New-Object System.Windows.Forms.Label; $labelReportTariff1.Text = "Tariff 1 (Base/Cost):"; $labelReportTariff1.Location = New-Object System.Drawing.Point(10, 25); $labelReportTariff1.AutoSize = $true; $labelReportTariff1.Font = $fontBold; $labelReportTariff1.ForeColor = $colorText
$listBoxReportTariff1 = New-Object System.Windows.Forms.ListBox; $listBoxReportTariff1.Location = New-Object System.Drawing.Point(10, 45); $listBoxReportTariff1.Size = New-Object System.Drawing.Size($reportsListBoxWidth, $reportsListBoxHeight); $listBoxReportTariff1.Font = $fontMono; $listBoxReportTariff1.IntegralHeight = $false; $listBoxReportTariff1.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelReportTariff2 = New-Object System.Windows.Forms.Label; $labelReportTariff2.Text = "Tariff 2 (Comparison):"; $labelReportTariff2.Location = New-Object System.Drawing.Point(280, 25); $labelReportTariff2.AutoSize = $true; $labelReportTariff2.Font = $fontBold; $labelReportTariff2.ForeColor = $colorText; $labelReportTariff2.Visible = $false 
$listBoxReportTariff2 = New-Object System.Windows.Forms.ListBox; $listBoxReportTariff2.Location = New-Object System.Drawing.Point(280, 45); $listBoxReportTariff2.Size = New-Object System.Drawing.Size($reportsListBoxWidth, $reportsListBoxHeight); $listBoxReportTariff2.Font = $fontMono; $listBoxReportTariff2.IntegralHeight = $false; $listBoxReportTariff2.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle; $listBoxReportTariff2.Visible = $false 
$groupBoxReportTariffSelect.Controls.AddRange(@($labelReportTariff1, $listBoxReportTariff1, $labelReportTariff2, $listBoxReportTariff2))

$labelSelectCsv = New-Object System.Windows.Forms.Label; $labelSelectCsv.Text = "Select Data File (CSV):"; $labelSelectCsv.AutoSize = $true; $labelSelectCsv.Font = $fontBold; $labelSelectCsv.ForeColor = $colorText
$textboxCsvPath = New-Object System.Windows.Forms.TextBox; $textboxCsvPath.Size = New-Object System.Drawing.Size($reportsInputWidth, 23); $textboxCsvPath.ReadOnly = $true; $textboxCsvPath.Font = $fontRegular; $textboxCsvPath.BackColor = $colorPanel
$buttonSelectCsv = New-Object System.Windows.Forms.Button; $buttonSelectCsv.Text = "Browse..."; $buttonSelectCsv.Size = New-Object System.Drawing.Size(80, 28); $buttonSelectCsv.Font = $fontRegular; $buttonSelectCsv.BackColor = $colorButtonBack; $buttonSelectCsv.ForeColor = $colorButtonFore; $buttonSelectCsv.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $buttonSelectCsv.FlatAppearance.BorderSize = 1; $buttonSelectCsv.FlatAppearance.BorderColor = $colorButtonBorder; $buttonSelectCsv.Enabled = $false

$groupBoxReportAspInput = New-Object System.Windows.Forms.GroupBox; $groupBoxReportAspInput.Text = "Desired Selling Price"; $groupBoxReportAspInput.Size = New-Object System.Drawing.Size(410, $groupBoxAspInputHeight); $groupBoxReportAspInput.Font = $fontRegular; $groupBoxReportAspInput.ForeColor = $colorText; $groupBoxReportAspInput.Visible = $false 
$labelDesiredAsp = New-Object System.Windows.Forms.Label; $labelDesiredAsp.Text = "Desired Avg Selling Price ($):"; $labelDesiredAsp.Location = New-Object System.Drawing.Point(10, 25); $labelDesiredAsp.AutoSize = $true; $labelDesiredAsp.ForeColor = $colorText
$textBoxDesiredAsp = New-Object System.Windows.Forms.TextBox; $textBoxDesiredAsp.Location = New-Object System.Drawing.Point(190, 22); $textBoxDesiredAsp.Size = New-Object System.Drawing.Size(100, 23); $textBoxDesiredAsp.Font = $fontRegular; $textBoxDesiredAsp.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$groupBoxReportAspInput.Controls.AddRange(@($labelDesiredAsp, $textBoxDesiredAsp))

$checkBoxApplyMargins = New-Object System.Windows.Forms.CheckBox; $checkBoxApplyMargins.Text = "Apply Calculated Margins to Tariff Files (Use with Caution!)"; 
$checkBoxApplyMargins.AutoSize = $true; $checkBoxApplyMargins.Font = $fontRegular; $checkBoxApplyMargins.ForeColor = [System.Drawing.Color]::DarkRed; $checkBoxApplyMargins.Visible = $false 

$buttonRunReport = New-Object System.Windows.Forms.Button; $buttonRunReport.Text = "Run Report"; 
$buttonRunReport.Size = New-Object System.Drawing.Size(120, 35); $buttonRunReport.Font = $fontBold; $buttonRunReport.BackColor = $colorButtonBack; $buttonRunReport.ForeColor = $colorButtonFore; $buttonRunReport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $buttonRunReport.FlatAppearance.BorderSize = 1; $buttonRunReport.FlatAppearance.BorderColor = $colorButtonBorder; $buttonRunReport.Enabled = $false

$textboxReportResults = New-Object System.Windows.Forms.TextBox; 
$textboxReportResults.Size = New-Object System.Drawing.Size(($reportsPanel.ClientSize.Width - $reportsPanelX - 15), 60); $textboxReportResults.Multiline = $true; $textboxReportResults.ReadOnly = $true; $textboxReportResults.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical; $textboxReportResults.Font = $fontMono; $textboxReportResults.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right; $textboxReportResults.BackColor = $colorPanel; $textboxReportResults.ForeColor = $colorText

$reportsPanel.Controls.AddRange(@(
    $labelSelectCustomer_Reports, $comboBoxSelectCustomer_Reports,
    $labelReportType, $comboBoxReportType,
    $groupBoxReportCarrierSelect,
    $groupBoxReportTariffSelect,
    $labelSelectCsv, $textboxCsvPath, $buttonSelectCsv,
    $groupBoxReportAspInput,
    $checkBoxApplyMargins,
    $buttonRunReport,
    $textboxReportResults
))

$mainForm.Controls.Add($statusBar)

# ============================================================
# Event Handlers Section
# ============================================================
$buttonLogin.Add_Click({ param($sender, $e); $username = $textboxUsername.Text; $password = $textboxPassword.Text; if (-not (Get-Command Test-PasswordHash -ErrorAction SilentlyContinue)) { $errorMsg = "FATAL ERROR: 'Test-PasswordHash' function not found."; Write-Error $errorMsg; [System.Windows.Forms.MessageBox]::Show($errorMsg, "Login Function Missing", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); return }; $script:currentUserProfile = $null; try { $script:currentUserProfile = Authenticate-User -Username $username -PasswordPlainText $password -UserAccountsFolderPath $script:userAccountsFolderPath } catch { $errorMsg = "Error during authentication: $($_.Exception.Message)"; Write-Error $errorMsg; [System.Windows.Forms.MessageBox]::Show($errorMsg, "Authentication Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) }; if ($script:currentUserProfile) { $statusBar.Text = "Login successful! Welcome, $($script:currentUserProfile.Username)."; $loginPanel.Visible = $false; $tabControlMain.Visible = $true; $comboBoxSelectCustomer_Quote.Items.Clear(); $comboBoxSelectCustomer_Settings.Items.Clear(); $comboBoxSelectCustomer_Reports.Items.Clear(); if ($script:allCustomerProfiles.Count -gt 0) { $customerNames = $script:allCustomerProfiles.Keys | Sort-Object; $comboBoxSelectCustomer_Quote.Items.AddRange($customerNames); $comboBoxSelectCustomer_Settings.Items.AddRange($customerNames); $comboBoxSelectCustomer_Reports.Items.AddRange($customerNames); $comboBoxSelectCustomer_Quote.SelectedIndex = 0; $comboBoxSelectCustomer_Settings.SelectedIndex = 0; $comboBoxSelectCustomer_Reports.SelectedIndex = 0; $comboBoxSelectCustomer_Quote.Enabled = $true; $comboBoxSelectCustomer_Settings.Enabled = $true; $comboBoxSelectCustomer_Reports.Enabled = $true; $script:selectedCustomerProfile = $script:allCustomerProfiles[$customerNames[0]]; Write-Host "DEBUG: Initial Selected Customer Profile set to '$($customerNames[0])'" } else { $comboBoxSelectCustomer_Quote.Items.Add("No Customers Found"); $comboBoxSelectCustomer_Settings.Items.Add("No Customers Found"); $comboBoxSelectCustomer_Reports.Items.Add("No Customers Found"); $comboBoxSelectCustomer_Quote.Enabled = $false; $comboBoxSelectCustomer_Settings.Enabled = $false; $comboBoxSelectCustomer_Reports.Enabled = $false; $script:selectedCustomerProfile = $null; Write-Warning "No customer profiles loaded from '$script:customerAccountsFolderPath'." }; $script:currentUserReportsFolder = Join-Path $script:reportsBaseFolderPath $script:currentUserProfile.Username; Ensure-DirectoryExists $script:currentUserReportsFolder; try { if ($script:selectedCustomerProfile) { Populate-TariffListBox -SelectedCarrier "Central" -ListBoxControl $listBoxTariffs -LabelControl $labelSelectedTariff -ButtonControl $buttonSetMargin -TextboxControl $textBoxNewMargin -CustomerProfile $script:selectedCustomerProfile } else { $statusBar.Text = "Ready. Please select a customer to view settings." } } catch { $statusBar.Text = "Error populating initial settings list: $($_.Exception.Message)" }; $textboxPassword.Clear(); $comboBoxReportType.Enabled = $true; $buttonSelectCsv.Enabled = $true; $buttonRunReport.Enabled = $true; $comboBoxReportType.SelectedIndex = 0; } else { if ($Error.Count -eq 0) { $statusBar.Text = "Login failed. Please check username and password."; [System.Windows.Forms.MessageBox]::Show("Login Failed. Please check username and password.", "Login Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning) }; $textboxPassword.Clear(); $textboxUsername.Focus() } })

$customerChangedHandler = { param($sender, $e); if ($sender.SelectedIndex -ge 0) { $selectedName = $sender.SelectedItem.ToString(); if ($script:allCustomerProfiles.ContainsKey($selectedName)) { $script:selectedCustomerProfile = $script:allCustomerProfiles[$selectedName]; Write-Host "DEBUG: Selected Customer Profile set to '$selectedName'"; $comboBoxes = @($comboBoxSelectCustomer_Quote, $comboBoxSelectCustomer_Settings, $comboBoxSelectCustomer_Reports); foreach($cb in $comboBoxes){ if($cb -ne $sender -and $cb.SelectedItem -ne $selectedName){ try {$cb.SelectedItem = $selectedName} catch {Write-Warning "Failed to sync customer selection to $($cb.Name)"} } }; if ($tabControlMain.SelectedTab -eq $tabPageSettings) { $selectedCarrierForSettings = "Central"; if ($radioSAIA.Checked) { $selectedCarrierForSettings = "SAIA" } elseif ($radioRL.Checked) { $selectedCarrierForSettings = "RL" }; try { Populate-TariffListBox -SelectedCarrier $selectedCarrierForSettings -ListBoxControl $listBoxTariffs -LabelControl $labelSelectedTariff -ButtonControl $buttonSetMargin -TextboxControl $textBoxNewMargin -CustomerProfile $script:selectedCustomerProfile } catch { $statusBar.Text = "Error refreshing settings list: $($_.Exception.Message)" } }; if ($tabControlMain.SelectedTab -eq $tabPageReports) { try { $comboBoxReportType_SelectedIndexChanged_ScriptBlock.Invoke($comboBoxReportType, [System.EventArgs]::Empty) } catch { $statusBar.Text = "Error refreshing report UI on customer change: $($_.Exception.Message)" } }; $statusBar.Text = "Selected Customer: $selectedName" } else { Write-Warning "Selected customer name '$selectedName' not found."; $script:selectedCustomerProfile = $null; $statusBar.Text = "Error selecting customer." } } else { $script:selectedCustomerProfile = $null; $statusBar.Text = "No customer selected." } }
$comboBoxSelectCustomer_Quote.Add_SelectedIndexChanged($customerChangedHandler); $comboBoxSelectCustomer_Settings.Add_SelectedIndexChanged($customerChangedHandler); $comboBoxSelectCustomer_Reports.Add_SelectedIndexChanged($customerChangedHandler)

$buttonGetQuote.Add_Click({ param($sender, $e); if ($null -eq $script:selectedCustomerProfile) { [System.Windows.Forms.MessageBox]::Show("Please select a customer first.", "Customer Not Selected", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return }; $textboxResults.Clear(); $textboxResults.Text = "Fetching rates for Customer: $($script:selectedCustomerProfile.CustomerName)... please wait."; $mainForm.Refresh(); $originZip = $textboxOriginZip.Text; $originCity = $textboxOriginCity.Text; $originState = $textboxOriginState.Text; $destZip = $textboxDestZip.Text; $destCity = $textboxDestCity.Text; $destState = $textboxDestState.Text; $weightInput = $textboxWeight.Text; $classInput = $textboxClass.Text; $lenInput = $textboxLength.Text; $widInput = $textboxWidthOpt.Text; $hgtInput = $textboxHeight.Text; $decValInput = $textboxDeclaredValue.Text; $validationErrors = [System.Collections.Generic.List[string]]::new(); if (-not ($originZip -match '^\d{5}$')) { $validationErrors.Add("Origin ZIP.") }; if ([string]::IsNullOrWhiteSpace($originCity)) { $validationErrors.Add("Origin City.") }; if (-not ($originState -match '^[A-Za-z]{2}$')) { $validationErrors.Add("Origin State.") }; if (-not ($destZip -match '^\d{5}$')) { $validationErrors.Add("Dest ZIP.") }; if ([string]::IsNullOrWhiteSpace($destCity)) { $validationErrors.Add("Dest City.") }; if (-not ($destState -match '^[A-Za-z]{2}$')) { $validationErrors.Add("Dest State.") }; $weight = $null; if ($weightInput -match '^\d+(\.\d+)?$' -and ([decimal]$weightInput -gt 0)) { $weight = [decimal]$weightInput } else { $validationErrors.Add("Weight.") }; $freightClass = $null; if ($classInput -match '^\d+(\.\d+)?$' -and ([double]$classInput -ge 50) -and ([double]$classInput -le 500) ) { $freightClass = $classInput } else { $validationErrors.Add("Class.") }; $itemLength = 1.0; if(-not [string]::IsNullOrWhiteSpace($lenInput)) { if($lenInput -match '^\d+(\.\d+)?$') { $itemLength = [float]$lenInput } else { $validationErrors.Add("Length.") } }; $itemWidth = 1.0; if(-not [string]::IsNullOrWhiteSpace($widInput)) { if($widInput -match '^\d+(\.\d+)?$') { $itemWidth = [float]$widInput } else { $validationErrors.Add("Width.") } }; $itemHeight = 1.0; if(-not [string]::IsNullOrWhiteSpace($hgtInput)) { if($hgtInput -match '^\d+(\.\d+)?$') { $itemHeight = [float]$hgtInput } else { $validationErrors.Add("Height.") } }; $declaredValue = 0.0; if(-not [string]::IsNullOrWhiteSpace($decValInput)) { if($decValInput -match '^\d+(\.\d+)?$') { $declaredValue = [decimal]$decValInput } else { $validationErrors.Add("Declared Value.") } }; if ($validationErrors.Count -gt 0) { $errorMsg = "Invalid Input:`n" + ($validationErrors -join "`n"); [System.Windows.Forms.MessageBox]::Show($errorMsg, "Validation Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); $textboxResults.Text = $errorMsg; return }; $statusBar.Text = "Fetching rates for Customer: $($script:selectedCustomerProfile.CustomerName), $originZip -> $destZip..."; $optionalShipmentDetails = [PSCustomObject]@{ OriginCity = $originCity; OriginState = $originState; DestinationCity = $destCity; DestinationState = $destState; ItemWidth = $itemLength; ItemHeight = $itemHeight; ItemLength = $itemWidth; DeclaredValue = $declaredValue; CustomerData = $null; QuoteType = 'Domestic'; }; $resultsText = [System.Text.StringBuilder]::new(); $resultsText.AppendLine("==================== SHIPMENT QUOTE ====================") | Out-Null; $resultsText.AppendLine("Quote Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')") | Out-Null; $resultsText.AppendLine("Broker User: $($script:currentUserProfile.Username)") | Out-Null; $resultsText.AppendLine("Customer:    $($script:selectedCustomerProfile.CustomerName)") | Out-Null; $resultsText.AppendLine("--------------------------------------------------------") | Out-Null; $resultsText.AppendLine("Origin:      $originCity, $originState $originZip") | Out-Null; $resultsText.AppendLine("Destination: $destCity, $destState $destZip") | Out-Null; $resultsText.AppendLine("Weight:      $($weight) lbs") | Out-Null; $resultsText.AppendLine("Class:       $($freightClass)") | Out-Null; if ($itemLength -ne 1.0 -or $itemWidth -ne 1.0 -or $itemHeight -ne 1.0) { $resultsText.AppendLine("Dimensions:  $itemLength"" L x $itemWidth"" W x $itemHeight"" H") | Out-Null }; if ($declaredValue -gt 0) { $resultsText.AppendLine("Declared Val:$($declaredValue.ToString("C"))") | Out-Null }; $resultsText.AppendLine("--------------------------------------------------------") | Out-Null; $resultsText.AppendLine("Carrier Options:") | Out-Null; $permittedCentralKeys = @{}; $permittedSAIAKeys = @{}; $permittedRLKeys = @{}; try { $permittedCentralKeys = Get-PermittedKeys -AllKeys $script:allCentralKeys -AllowedKeyNames $script:selectedCustomerProfile.AllowedCentralKeys; $permittedSAIAKeys = Get-PermittedKeys -AllKeys $script:allSAIAKeys -AllowedKeyNames $script:selectedCustomerProfile.AllowedSAIAKeys; $permittedRLKeys = Get-PermittedKeys -AllKeys $script:allRLKeys -AllowedKeyNames $script:selectedCustomerProfile.AllowedRLKeys } catch { [System.Windows.Forms.MessageBox]::Show("Error getting permitted keys for customer: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); $textboxResults.Text = "Error getting permitted keys for customer."; return }; $quoteTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; $finalQuotes = @(); $CurrentVerbosePreference = $VerbosePreference; $VerbosePreference = 'SilentlyContinue'; try { $centralRates = @{}; if ($permittedCentralKeys.Count -gt 0) { $textboxResults.AppendText("Querying Central Transport..." + "`r`n"); $mainForm.Refresh(); foreach ($tariffFileName in ($permittedCentralKeys.Keys | Sort-Object)) { try { $keyData = $permittedCentralKeys[$tariffFileName]; if (-not $keyData.ContainsKey('accessCode') -or -not $keyData.ContainsKey('customerNumber')) { throw "'accessCode'/'customerNumber' missing." }; $cost = Invoke-CentralTransportApi -ApiKey $keyData.accessCode -OriginZip $originZip -DestinationZip $destZip -Weight $weight -FreightClass $freightClass -customerNumber $keyData.customerNumber; if ($cost -ne $null) { $centralRates[$tariffFileName] = $cost } } catch { $resultsText.AppendLine("  Central ($tariffFileName): Error - $($_.Exception.Message)") | Out-Null } }; $centralLowest = Get-MinimumRate -RateResults $centralRates; if ($centralLowest -ne $null) { $lowestTariffData = $permittedCentralKeys[$centralLowest.TariffName]; $marginToUse = $Global:DefaultMarginPercentage; if ($lowestTariffData -and $lowestTariffData.ContainsKey('MarginPercent')) { try { $marginToUse = [double]$lowestTariffData.MarginPercent } catch {} }; $centralQuoteDetails = Calculate-QuotePrice -LowestCarrierCost $centralLowest.Cost -OriginZip $originZip -DestinationZip $destZip -Weight $weight -FreightClass $freightClass -MarginPercent $marginToUse; if ($centralQuoteDetails.FinalPrice -ne $null) { $resultsText.AppendLine(("  {0,-20}: {1,-15} (Tariff: {2})" -f "Central Transport", $centralQuoteDetails.FinalPrice.ToString("C"), $centralLowest.TariffName)) | Out-Null; $finalQuotes += [PSCustomObject]@{Carrier='Central Transport'; Tariff=$centralLowest.TariffName; Price=$centralQuoteDetails.FinalPrice; Cost=$centralLowest.Cost}; Write-QuoteToHistory -Carrier 'Central Transport' -TariffName $centralLowest.TariffName -OriginZip $originZip -DestinationZip $destZip -Weight $weight -FreightClass $freightClass -LowestCost $centralLowest.Cost -FinalQuotedPrice $centralQuoteDetails.FinalPrice -QuoteTimestamp $quoteTimestamp } else { $resultsText.AppendLine("  Central Transport    : Error calculating final price.") | Out-Null } } else { $resultsText.AppendLine("  Central Transport    : No valid rates found.") | Out-Null } } else { $resultsText.AppendLine("  Central Transport    : No permitted keys found for this customer.") | Out-Null }; $saiaRates = @{}; if ($permittedSAIAKeys.Count -gt 0) { $textboxResults.AppendText("Querying SAIA..." + "`r`n"); $mainForm.Refresh(); foreach ($tariffFileName in ($permittedSAIAKeys.Keys | Sort-Object)) { try { $keyData = $permittedSAIAKeys[$tariffFileName]; $cost = Invoke-SAIAApi -OriginZip $originZip -DestinationZip $destZip -OriginCity $originCity -OriginState $originState -DestinationCity $destCity -DestinationState $destState -Weight $weight -Class $freightClass -KeyData $keyData; if ($cost -ne $null) { $saiaRates[$tariffFileName] = $cost } } catch { $resultsText.AppendLine("  SAIA ($tariffFileName): Error - $($_.Exception.Message)") | Out-Null } }; $saiaLowest = Get-MinimumRate -RateResults $saiaRates; if ($saiaLowest -ne $null) { $lowestTariffData = $permittedSAIAKeys[$saiaLowest.TariffName]; $marginToUse = $Global:DefaultMarginPercentage; if ($lowestTariffData -and $lowestTariffData.ContainsKey('MarginPercent')) { try { $marginToUse = [double]$lowestTariffData.MarginPercent } catch {} }; $saiaQuoteDetails = Calculate-QuotePrice -LowestCarrierCost $saiaLowest.Cost -OriginZip $originZip -DestinationZip $destZip -Weight $weight -FreightClass $freightClass -MarginPercent $marginToUse; if ($saiaQuoteDetails.FinalPrice -ne $null) { $resultsText.AppendLine(("  {0,-20}: {1,-15} (Tariff: {2})" -f "SAIA", $saiaQuoteDetails.FinalPrice.ToString("C"), $saiaLowest.TariffName)) | Out-Null; $finalQuotes += [PSCustomObject]@{Carrier='SAIA'; Tariff=$saiaLowest.TariffName; Price=$saiaQuoteDetails.FinalPrice; Cost=$saiaLowest.Cost}; Write-QuoteToHistory -Carrier 'SAIA' -TariffName $saiaLowest.TariffName -OriginZip $originZip -DestinationZip $destZip -Weight $weight -FreightClass $freightClass -LowestCost $saiaLowest.Cost -FinalQuotedPrice $saiaQuoteDetails.FinalPrice -QuoteTimestamp $quoteTimestamp } else { $resultsText.AppendLine("  SAIA                 : Error calculating final price.") | Out-Null } } else { $resultsText.AppendLine("  SAIA                 : No valid rates found.") | Out-Null } } else { $resultsText.AppendLine("  SAIA                 : No permitted keys found for this customer.") | Out-Null }; $rlRates = @{}; if ($permittedRLKeys.Count -gt 0) { $textboxResults.AppendText("Querying R+L Carriers..." + "`r`n"); $mainForm.Refresh(); foreach ($tariffFileName in ($permittedRLKeys.Keys | Sort-Object)) { try { $keyData = $permittedRLKeys[$tariffFileName]; if (-not $keyData.ContainsKey('APIKey')) { throw "'APIKey' missing." }; $cost = Invoke-RLApi -OriginZip $originZip -DestinationZip $destZip -Weight $weight -Class $freightClass -KeyData $keyData -ShipmentDetails $optionalShipmentDetails; if ($cost -ne $null) { $rlRates[$tariffFileName] = $cost } } catch { $resultsText.AppendLine("  R+L ($tariffFileName): Error - $($_.Exception.Message)") | Out-Null } }; $rlLowest = Get-MinimumRate -RateResults $rlRates; if ($rlLowest -ne $null) { $lowestTariffData = $permittedRLKeys[$rlLowest.TariffName]; $marginToUse = $Global:DefaultMarginPercentage; if ($lowestTariffData -and $lowestTariffData.ContainsKey('MarginPercent')) { try { $marginToUse = [double]$lowestTariffData.MarginPercent } catch {} }; $rlQuoteDetails = Calculate-QuotePrice -LowestCarrierCost $rlLowest.Cost -OriginZip $originZip -DestinationZip $destZip -Weight $weight -FreightClass $freightClass -MarginPercent $marginToUse; if ($rlQuoteDetails.FinalPrice -ne $null) { $resultsText.AppendLine(("  {0,-20}: {1,-15} (Tariff: {2})" -f "R+L Carriers", $rlQuoteDetails.FinalPrice.ToString("C"), $rlLowest.TariffName)) | Out-Null; $finalQuotes += [PSCustomObject]@{Carrier='R+L Carriers'; Tariff=$rlLowest.TariffName; Price=$rlQuoteDetails.FinalPrice; Cost=$rlLowest.Cost}; Write-QuoteToHistory -Carrier 'R+L Carriers' -TariffName $rlLowest.TariffName -OriginZip $originZip -DestinationZip $destZip -Weight $weight -FreightClass $freightClass -LowestCost $rlLowest.Cost -FinalQuotedPrice $rlQuoteDetails.FinalPrice -QuoteTimestamp $quoteTimestamp } else { $resultsText.AppendLine("  R+L Carriers         : Error calculating final price.") | Out-Null } } else { $resultsText.AppendLine("  R+L Carriers         : No valid rates found.") | Out-Null } } else { $resultsText.AppendLine("  R+L Carriers         : No permitted keys found for this customer.") | Out-Null } } finally { $VerbosePreference = $CurrentVerbosePreference }; $resultsText.AppendLine("========================================================") | Out-Null; $resultsText.AppendLine("* Prices are estimates and subject to verification. *") | Out-Null; $resultsText.AppendLine("--- End of Quote ---") | Out-Null; $textboxResults.Text = $resultsText.ToString(); $statusBar.Text = "Quote generation complete for Customer: $($script:selectedCustomerProfile.CustomerName)." })

# --- MODIFICATION: Store the script block for re-use and dynamic Y-positioning ---
$comboBoxReportType_SelectedIndexChanged_ScriptBlock = {
    param($sender, $e) # Standard event handler parameters
    $selectedReport = $comboBoxReportType.SelectedItem.ToString()
    
    # Base Y position after the static controls (Customer and Report Type dropdowns)
    $dynamicY = $reportsPanelY + (2 * $reportsControlSpacing) 

    # Default visibility states
    $groupBoxReportCarrierSelect.Visible = $false
    $groupBoxReportTariffSelect.Visible = $false
    $labelReportTariff2.Visible = $false
    $listBoxReportTariff2.Visible = $false
    $groupBoxReportAspInput.Visible = $false
    $checkBoxApplyMargins.Visible = $false

    # Logic for single-carrier specific UI elements
    if ($selectedReport -eq "Carrier Comparison" -or $selectedReport -eq "Avg Required Margin" -or $selectedReport -eq "Required Margin for ASP") {
        $groupBoxReportCarrierSelect.Location = New-Object System.Drawing.Point($reportsPanelX, $dynamicY)
        $groupBoxReportCarrierSelect.Visible = $true
        $dynamicY += $groupBoxReportCarrierSelect.Height + $verticalPaddingBetweenControls
        
        $groupBoxReportTariffSelect.Location = New-Object System.Drawing.Point($reportsPanelX, $dynamicY)
        $groupBoxReportTariffSelect.Visible = $true
        $dynamicY += $groupBoxReportTariffSelect.Height + $verticalPaddingBetweenControls
        
        if ($selectedReport -eq "Carrier Comparison" -or $selectedReport -eq "Avg Required Margin") {
            $labelReportTariff2.Visible = $true
            $listBoxReportTariff2.Visible = $true
            $labelReportTariff1.Text = "Tariff 1 (Base):"
        } else { # This is "Required Margin for ASP"
            $labelReportTariff1.Text = "Select Tariff:"
        }
    }
    # If not a single carrier report, $dynamicY remains as is (after Customer and Report Type dropdowns)

    # Position CSV controls
    $labelSelectCsv.Location = New-Object System.Drawing.Point($reportsPanelX, $dynamicY)
    $textboxCsvPath.Location = New-Object System.Drawing.Point(($reportsPanelX + $reportsLabelWidth + 5), ($dynamicY - 3))
    $buttonSelectCsv.Location = New-Object System.Drawing.Point(($reportsPanelX + $reportsLabelWidth + $reportsInputWidth + 10), ($dynamicY - 5))
    $dynamicY += $reportsControlSpacing 

    # Visibility and positioning logic for ASP input and Apply Margins checkbox
    if ($selectedReport -eq "Required Margin for ASP") {
        $groupBoxReportAspInput.Location = New-Object System.Drawing.Point($reportsPanelX, $dynamicY)
        $groupBoxReportAspInput.Visible = $true
        $checkBoxApplyMargins.Visible = $false 
        $dynamicY += $groupBoxReportAspInput.Height + $verticalPaddingBetweenControls
    } elseif ($selectedReport -eq "Cross-Carrier ASP Analysis") {
        $groupBoxReportAspInput.Location = New-Object System.Drawing.Point($reportsPanelX, $dynamicY)
        $groupBoxReportAspInput.Visible = $true
        $dynamicY += $groupBoxReportAspInput.Height + $verticalPaddingBetweenControls
        
        $checkBoxApplyMargins.Location = New-Object System.Drawing.Point($reportsPanelX, $dynamicY)
        $checkBoxApplyMargins.Visible = $true
        $dynamicY += $checkBoxApplyMargins.Height + $verticalPaddingBetweenControls
    } elseif ($selectedReport -eq "Margins by History") {
        $groupBoxReportAspInput.Visible = $false 
        
        $checkBoxApplyMargins.Location = New-Object System.Drawing.Point($reportsPanelX, $dynamicY) 
        $checkBoxApplyMargins.Visible = $true
        $dynamicY += $checkBoxApplyMargins.Height + $verticalPaddingBetweenControls
    } else { 
        $groupBoxReportAspInput.Visible = $false
        $checkBoxApplyMargins.Visible = $false
    }
    
    $buttonRunReport.Location = New-Object System.Drawing.Point(($reportsPanelX + $reportsLabelWidth + 5), $dynamicY)
    $dynamicY += $buttonRunReport.Height + $verticalPaddingBetweenControls
    $textboxReportResults.Location = New-Object System.Drawing.Point($reportsPanelX, $dynamicY)

    if ($groupBoxReportCarrierSelect.Visible -and $script:selectedCustomerProfile) { 
        $selectedCarrierForReports = "Central"
        if ($radioReportSAIA.Checked) { $selectedCarrierForReports = "SAIA" } 
        elseif ($radioReportRL.Checked) { $selectedCarrierForReports = "RL" }
        try { 
            Populate-ReportTariffListBoxes -SelectedCarrier $selectedCarrierForReports -ReportType $selectedReport -CustomerProfile $script:selectedCustomerProfile -ListBox1 $listBoxReportTariff1 -Label1 $labelReportTariff1 -ListBox2 $listBoxReportTariff2 -Label2 $labelReportTariff2 
        } catch { 
            $statusBar.Text = "Error refreshing report tariff list: $($_.Exception.Message)" 
        } 
    } 
}
$comboBoxReportType.Add_SelectedIndexChanged($comboBoxReportType_SelectedIndexChanged_ScriptBlock)

$reportCarrierChangedHandler = { param($sender, $e); if ($sender.Checked -and $script:selectedCustomerProfile) { $selectedCarrierForReports = "Central"; if ($radioReportSAIA.Checked) { $selectedCarrierForReports = "SAIA" } elseif ($radioReportRL.Checked) { $selectedCarrierForReports = "RL" }; $selectedReport = $comboBoxReportType.SelectedItem.ToString(); try { Populate-ReportTariffListBoxes -SelectedCarrier $selectedCarrierForReports -ReportType $selectedReport -CustomerProfile $script:selectedCustomerProfile -ListBox1 $listBoxReportTariff1 -Label1 $labelReportTariff1 -ListBox2 $listBoxReportTariff2 -Label2 $labelReportTariff2 } catch { $statusBar.Text = "Error refreshing report tariff list: $($_.Exception.Message)" } } }
$radioReportCentral.Add_CheckedChanged($reportCarrierChangedHandler); $radioReportSAIA.Add_CheckedChanged($reportCarrierChangedHandler); $radioReportRL.Add_CheckedChanged($reportCarrierChangedHandler)
$buttonSelectCsv.Add_Click({ param($sender, $e); $csvPath = Select-CsvFile -DialogTitle "Select Report Data CSV" -InitialDirectory $script:shipmentDataFolderPath; if (-not [string]::IsNullOrWhiteSpace($csvPath)) { $textboxCsvPath.Text = $csvPath } })
$buttonRunReport.Add_Click({ param($sender, $e); $textboxReportResults.Clear(); $statusBar.Text = "Starting report generation..."; $mainForm.Refresh(); if ($null -eq $script:selectedCustomerProfile) { [System.Windows.Forms.MessageBox]::Show("Please select a customer.", "Input Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return }; if ($comboBoxReportType.SelectedIndex -lt 0) { [System.Windows.Forms.MessageBox]::Show("Please select a report type.", "Input Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return }; if ([string]::IsNullOrWhiteSpace($textboxCsvPath.Text) -or -not (Test-Path $textboxCsvPath.Text -PathType Leaf)) { [System.Windows.Forms.MessageBox]::Show("Please select a valid CSV data file.", "Input Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning); return }; $selectedReportType = $comboBoxReportType.SelectedItem.ToString(); $csvPath = $textboxCsvPath.Text; $customerProfile = $script:selectedCustomerProfile; $brokerProfile = $script:currentUserProfile; $reportsFolder = $script:currentUserReportsFolder; $reportFunction = $null; $reportParams = @{ CsvFilePath = $csvPath; UserReportsFolder = $reportsFolder };  if ($brokerProfile -and $brokerProfile.PSObject.Properties.Name -contains 'Username') {$reportParams.Username = $brokerProfile.Username} else {$reportParams.Username = "UnknownUser"} ; $selectedCarrier = $null; $key1Data = $null; $key2Data = $null; try { if ($groupBoxReportCarrierSelect.Visible) { if ($radioReportCentral.Checked) { $selectedCarrier = "Central" } elseif ($radioReportSAIA.Checked) { $selectedCarrier = "SAIA" } elseif ($radioReportRL.Checked) { $selectedCarrier = "RL" } else { throw "No carrier selected." }; if ($listBoxReportTariff1.SelectedIndex -lt 0) { throw "Please select Tariff 1." }; $tariff1Name = $listBoxReportTariff1.SelectedItem.ToString(); $allKeys = switch($selectedCarrier){"Central"{$script:allCentralKeys}"SAIA"{$script:allSAIAKeys}"RL"{$script:allRLKeys}}; if (-not $allKeys.ContainsKey($tariff1Name)) { throw "Selected Tariff 1 ('$tariff1Name') data not found." }; $key1Data = $allKeys[$tariff1Name]; if ($listBoxReportTariff2.Visible) { if ($listBoxReportTariff2.SelectedIndex -lt 0) { throw "Please select Tariff 2." }; $tariff2Name = $listBoxReportTariff2.SelectedItem.ToString(); if ($tariff1Name -eq $tariff2Name) { throw "Tariff 1 and Tariff 2 cannot be the same."}; if (-not $allKeys.ContainsKey($tariff2Name)) { throw "Selected Tariff 2 ('$tariff2Name') data not found." }; $key2Data = $allKeys[$tariff2Name] } }; switch ($selectedReportType) { "Carrier Comparison" { $reportFunction = "Run-{0}ComparisonReportGUI" -f $selectedCarrier; $reportParams.Key1Data = $key1Data; $reportParams.Key2Data = $key2Data } "Avg Required Margin" { $reportFunction = "Run-{0}MarginReportGUI" -f $selectedCarrier; $reportParams.BaseKeyData = $key1Data; $reportParams.ComparisonKeyData = $key2Data } "Required Margin for ASP" { $reportFunction = "Calculate-{0}MarginForASPReportGUI" -f $selectedCarrier; if (-not $groupBoxReportAspInput.Visible) { Write-Warning "ASP Input groupbox is not visible for 'Required Margin for ASP' report. This might be a UI logic error."; throw "ASP Input not visible for 'Required Margin for ASP' report type?" }; $aspInput = $textBoxDesiredAsp.Text; if (-not ($aspInput -match '^\d+(\.\d+)?$' -and ([decimal]$aspInput -gt 0))) { throw "Invalid Desired ASP value." }; $reportParams.CostAccountInfo = $key1Data; $reportParams.DesiredASP = [decimal]$aspInput } "Cross-Carrier ASP Analysis" { $reportFunction = "Run-CrossCarrierASPAnalysisGUI"; if (-not $groupBoxReportAspInput.Visible) { Write-Warning "ASP Input groupbox is not visible for 'Cross-Carrier ASP Analysis' report. This might be a UI logic error."; throw "ASP Input not visible for 'Cross-Carrier ASP Analysis' report type?" }; $aspInput = $textBoxDesiredAsp.Text; if (-not ($aspInput -match '^\d+(\.\d+)?$' -and ([decimal]$aspInput -gt 0))) { throw "Invalid Desired ASP value." }; $reportParams.Remove('Username'); $reportParams.BrokerProfile = $brokerProfile; $reportParams.SelectedCustomerProfile = $customerProfile; $reportParams.AllCentralKeys = $script:allCentralKeys; $reportParams.AllSAIAKeys = $script:allSAIAKeys; $reportParams.AllRLKeys = $script:allRLKeys; $reportParams.DesiredASPValue = [decimal]$aspInput; $reportParams.ApplyMargins = $checkBoxApplyMargins.Checked; $reportParams.ASPFromHistory = $false } "Margins by History" { $reportFunction = "Run-MarginsByHistoryAnalysisGUI"; $reportParams = @{ BrokerProfile = $brokerProfile; SelectedCustomerProfile = $customerProfile; ReportsBaseFolder = $reportsFolder; UserReportsFolder = $reportsFolder; AllCentralKeys = $script:allCentralKeys; AllSAIAKeys = $script:allSAIAKeys; AllRLKeys = $script:allRLKeys; CsvFilePath = $csvPath; ApplyMargins = $checkBoxApplyMargins.Checked } } default { throw "Selected report type '$selectedReportType' is not implemented." } }; $statusBar.Text = "Running Report: $selectedReportType..."; $mainForm.Refresh(); Write-Host "DEBUG: Calling $reportFunction with params:"; $reportParams | Format-List | Out-String | Write-Host; $reportOutputPath = & $reportFunction @reportParams; if ($reportOutputPath -and (Test-Path $reportOutputPath)) { $textboxReportResults.Text = "Report generated successfully!`nPath: $reportOutputPath"; $statusBar.Text = "Report Complete."; $result = [System.Windows.Forms.MessageBox]::Show("Report generated successfully.`n`nOpen the report file now?`n$reportOutputPath", "Report Complete", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information); if ($result -eq [System.Windows.Forms.DialogResult]::Yes) { try { Open-FileExplorer -Path $reportOutputPath } catch { [System.Windows.Forms.MessageBox]::Show("Failed to open report file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) } } } else { $textboxReportResults.Text = "Report generation failed. Check console output for details."; $statusBar.Text = "Report Failed."; [System.Windows.Forms.MessageBox]::Show("Report generation failed. Check console output.", "Report Failed", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) } } catch { $errorMsg = "Error running report '$selectedReportType':`n$($_.Exception.Message)"; Write-Error $errorMsg; $textboxReportResults.Text = $errorMsg; $statusBar.Text = "Report Error."; [System.Windows.Forms.MessageBox]::Show($errorMsg, "Report Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error) } })

$mainForm.Add_Shown({ if ($loginPanel.Visible) { $textboxUsername.Focus() } })
[void]$mainForm.ShowDialog()
Write-Host "GUI Closed."
