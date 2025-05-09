<#
.SYNOPSIS
GUI Front-end for the Transportation Management System Tool.
Uses Windows Forms and leverages existing TMS PowerShell modules.
Broker logs in (from user_accounts), then selects a customer (from customer_accounts) to work with.
Includes Quote, Settings, and Reports tabs. Handles multiple commodity items.
Integrates AAA Cooper Transportation.
#>

# --- Load Required Assemblies ---
try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
} catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to load required .NET Assemblies. Ensure .NET Framework is available.", "Fatal Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
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
    $moduleFiles = @(
        "TMS_Config.ps1",
        "TMS_Helpers_General.ps1",
        "TMS_GUI_Helpers.ps1",
        "TMS_Auth.ps1",
        "TMS_Helpers_Central.ps1",
        "TMS_Helpers_SAIA.ps1",
        "TMS_Helpers_RL.ps1",
        "TMS_Helpers_Averitt.ps1",
        "TMS_Helpers_AAACooper.ps1", 
        "TMS_Carrier_Central.ps1",
        "TMS_Carrier_SAIA.ps1",
        "TMS_Carrier_RL.ps1",
        "TMS_Carrier_Averitt.ps1",
        "TMS_Carrier_AAACooper.ps1", 
        "TMS_Reports.ps1",
        "TMS_Settings.ps1"
    )

    foreach ($moduleFile in $moduleFiles) {
        if ($moduleFile.StartsWith('#')) { continue }
        $modulePath = Join-Path $script:scriptRoot $moduleFile
        Write-Verbose "Attempting to load module '$moduleFile' from '$modulePath'"
        if (-not (Test-Path $modulePath -PathType Leaf)) {
             if ($moduleFile -like "*Averitt*" -or $moduleFile -like "*AAACooper*") { 
                 Write-Warning "Optional module not found: '$moduleFile'. Related functionality may be limited."
                 continue
             } else {
                throw "Required module not found: '$moduleFile' at '$modulePath'"
             }
        }
        try {
            . $modulePath
        } catch { Write-Error "ERROR loading module '$moduleFile': $($_.Exception.Message)"; throw $_ }
    }
    # --- Final Verification AFTER loop ---
     if (-not (Get-Command Populate-TariffListBox -ErrorAction SilentlyContinue)) { throw "Required function 'Populate-TariffListBox' not found."}
     if (-not (Get-Command Populate-ReportTariffListBoxes -ErrorAction SilentlyContinue)) { throw "Required function 'Populate-ReportTariffListBoxes' not found."}
     if (-not (Get-Command Update-TariffMargin -ErrorAction SilentlyContinue)) { throw "Required function 'Update-TariffMargin' not found."}
     # Verify API helpers exist (allow optional ones to be missing)
     if (-not (Get-Command Invoke-CentralTransportApi -EA SilentlyContinue)){throw "Invoke-CentralTransportApi missing."}
     if (-not (Get-Command Invoke-SAIAApi -EA SilentlyContinue)){throw "Invoke-SAIAApi missing."}
     if (-not (Get-Command Invoke-RLApi -EA SilentlyContinue)){throw "Invoke-RLApi missing."}
     if (-not (Get-Command Invoke-AverittApi -EA SilentlyContinue)){Write-Warning "Invoke-AverittApi missing. Averitt quotes may be disabled."}
     if (-not (Get-Command Invoke-AAACooperApi -EA SilentlyContinue)){Write-Warning "Invoke-AAACooperApi missing. AAA Cooper quotes may be disabled."}


    Write-Verbose "Module loading complete."
} catch {
     $errorMessage = "FATAL: Failed to load a module or required function. GUI cannot start.`nError: $($_.Exception.Message)"; Write-Error $errorMessage; [System.Windows.Forms.MessageBox]::Show($errorMessage, "Module Load Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error); Exit 1
}

# --- Resolve Full Paths for Data Folders ---
$CentralKeysFolderName = $script:defaultCentralKeysFolderName
$SAIAKeysFolderName = $script:defaultSAIAKeysFolderName
$RLKeysFolderName = $script:defaultRLKeysFolderName
$AverittKeysFolderName = $script:defaultAverittKeysFolderName 
$AAACooperKeysFolderName = $script:defaultAAACooperKeysFolderName 

$UserAccountsFolderName = $script:defaultUserAccountsFolderName
$CustomerAccountsFolderName = $script:defaultCustomerAccountsFolderName
$ReportsBaseFolderName = $script:defaultReportsBaseFolderName
$ShipmentDataFolderName = $script:defaultShipmentDataFolderName

$script:centralKeysFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $CentralKeysFolderName
$script:saiaKeysFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $SAIAKeysFolderName
$script:rlKeysFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $RLKeysFolderName
$script:averittKeysFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $AverittKeysFolderName
$script:aaaCooperKeysFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $AAACooperKeysFolderName 

$script:userAccountsFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $UserAccountsFolderName
$script:customerAccountsFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $CustomerAccountsFolderName
$script:reportsBaseFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $ReportsBaseFolderName
$script:shipmentDataFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $ShipmentDataFolderName

# --- Ensure Required Base Folders Exist ---
Write-Verbose "Ensuring base data directories exist..."
try {
    @($script:centralKeysFolderPath, $script:saiaKeysFolderPath, $script:rlKeysFolderPath, $script:averittKeysFolderPath, $script:aaaCooperKeysFolderPath, 
      $script:userAccountsFolderPath, $script:customerAccountsFolderPath, $script:reportsBaseFolderPath, $script:shipmentDataFolderPath) | ForEach-Object { Ensure-DirectoryExists -Path $_ }
}
catch { $errorMessage = "FATAL: Failed to ensure required data directories exist.`nError: $($_.Exception.Message)"; Write-Error $errorMessage; [System.Windows.Forms.MessageBox]::Show($errorMessage, "Directory Error"); Exit 1 }
Write-Verbose "Base directory check complete."

# --- Pre-load All Carrier Keys/Data ---
Write-Verbose "Loading all available carrier keys/accounts/margins..."
$script:allCentralKeys = @{}; $script:allSAIAKeys = @{}; $script:allRLKeys = @{}; $script:allAverittKeys = @{}; $script:allAAACooperKeys = @{} 
try {
    $script:allCentralKeys = Load-KeysFromFolder -KeysFolderPath $script:centralKeysFolderPath -CarrierName "Central Transport"
    $script:allSAIAKeys = Load-KeysFromFolder -KeysFolderPath $script:saiaKeysFolderPath -CarrierName "SAIA"
    $script:allRLKeys = Load-KeysFromFolder -KeysFolderPath $script:rlKeysFolderPath -CarrierName "RL Carriers"
    $script:allAverittKeys = Load-KeysFromFolder -KeysFolderPath $script:averittKeysFolderPath -CarrierName "Averitt"
    $script:allAAACooperKeys = Load-KeysFromFolder -KeysFolderPath $script:aaaCooperKeysFolderPath -CarrierName "AAACooper" 
} catch {
    $loadErrorMsg = "ERROR during Load-KeysFromFolder: $($_.Exception.Message). Check paths."; Write-Error $loadErrorMsg
    [System.Windows.Forms.MessageBox]::Show($loadErrorMsg, "Key Loading Warning", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
    # Initialize to empty if loading fails to prevent null reference errors later
    $script:allCentralKeys = if($null -eq $script:allCentralKeys){@()}else{$script:allCentralKeys}
    $script:allSAIAKeys = if($null -eq $script:allSAIAKeys){@()}else{$script:allSAIAKeys}
    $script:allRLKeys = if($null -eq $script:allRLKeys){@()}else{$script:allRLKeys}
    $script:allAverittKeys = if($null -eq $script:allAverittKeys){@()}else{$script:allAverittKeys}
    $script:allAAACooperKeys = if($null -eq $script:allAAACooperKeys){@()}else{$script:allAAACooperKeys}
}
Write-Verbose "Key/Account/Margin loading complete."

# --- Pre-load All Customer Profiles ---
Write-Verbose "Loading all available customer profiles..."
$script:allCustomerProfiles = Load-AllCustomerProfiles -UserAccountsFolderPath $script:customerAccountsFolderPath
Write-Verbose "Customer profile loading complete."
if($script:allCustomerProfiles.Count -eq 0){ Write-Warning "No customer profiles loaded from '$($script:customerAccountsFolderPath)'!"}
else { Write-Verbose "$($script:allCustomerProfiles.Count) customer profiles loaded." }


# --- Script Variables ---
$script:currentUserProfile = $null       # Hashtable for the logged-in broker
$script:selectedCustomerProfile = $null # Hashtable for the selected customer (Mainly for Quote tab direct use)
$script:currentUserReportsFolder = $null # Path for reports

# --- UI Styles and Fonts ---
$fontRegular = New-Object System.Drawing.Font("Segoe UI", 9); $fontBold = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold); $fontTitle = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold); $fontMono = New-Object System.Drawing.Font("Consolas", 9)
$colorBackground = [System.Drawing.Color]::FromArgb(240, 242, 245); $colorPanel = [System.Drawing.Color]::White; $colorPrimary = [System.Drawing.Color]::FromArgb(0, 120, 215); $colorButtonBack = $colorPrimary; $colorButtonFore = [System.Drawing.Color]::White; $colorButtonBorder = [System.Drawing.Color]::FromArgb(0, 90, 180); $colorText = [System.Drawing.Color]::FromArgb(30, 30, 30); $colorTextLight = [System.Drawing.Color]::FromArgb(100, 100, 100);
$paddingMedium = New-Object System.Windows.Forms.Padding(10)

# --- Main Form ---
$mainForm = New-Object System.Windows.Forms.Form; $mainForm.Text = "TMS GUI Tool (Broker Mode)"; $mainForm.Size = New-Object System.Drawing.Size(850, 700); $mainForm.MinimumSize = New-Object System.Drawing.Size(800, 600); $mainForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen; $mainForm.MaximizeBox = $true; $mainForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::Sizable; $mainForm.BackColor = $colorBackground; $mainForm.Font = $fontRegular

# --- Login Panel ---
$loginPanel = New-Object System.Windows.Forms.Panel; $loginPanel.Location = New-Object System.Drawing.Point(10, 10); $loginPanel.Size = New-Object System.Drawing.Size(320, 170); $loginPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle; $loginPanel.Anchor = [System.Windows.Forms.AnchorStyles]::None; $loginPanel.BackColor = $colorPanel; $loginPanel.Padding = $paddingMedium
$loginTitleLabel = New-Object System.Windows.Forms.Label; $loginTitleLabel.Text = "Broker Login"; $loginTitleLabel.Font = $fontTitle; $loginTitleLabel.ForeColor = $colorPrimary; $loginTitleLabel.AutoSize = $true; $loginTitleLabel.Location = New-Object System.Drawing.Point(10, 10)
$labelUsername = New-Object System.Windows.Forms.Label; $labelUsername.Text = "Username:"; $labelUsername.Location = New-Object System.Drawing.Point(10, 55); $labelUsername.AutoSize = $true; $labelUsername.ForeColor = $colorText
$textboxUsername = New-Object System.Windows.Forms.TextBox; $textboxUsername.Location = New-Object System.Drawing.Point(95, 52); $textboxUsername.Size = New-Object System.Drawing.Size(200, 23); $textboxUsername.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelPassword = New-Object System.Windows.Forms.Label; $labelPassword.Text = "Password:"; $labelPassword.Location = New-Object System.Drawing.Point(10, 88); $labelPassword.AutoSize = $true; $labelPassword.ForeColor = $colorText
$textboxPassword = New-Object System.Windows.Forms.TextBox; $textboxPassword.Location = New-Object System.Drawing.Point(95, 85); $textboxPassword.Size = New-Object System.Drawing.Size(200, 23); $textboxPassword.UseSystemPasswordChar = $true; $textboxPassword.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$buttonLogin = New-Object System.Windows.Forms.Button; $buttonLogin.Text = "Login"; $buttonLogin.Location = New-Object System.Drawing.Point(120, 125); $buttonLogin.Size = New-Object System.Drawing.Size(80, 30); $buttonLogin.Font = $fontBold; $buttonLogin.BackColor = $colorButtonBack; $buttonLogin.ForeColor = $colorButtonFore; $buttonLogin.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $buttonLogin.FlatAppearance.BorderSize = 1; $buttonLogin.FlatAppearance.BorderColor = $colorButtonBorder
$loginPanel.Controls.AddRange(@($loginTitleLabel, $labelUsername, $textboxUsername, $labelPassword, $textboxPassword, $buttonLogin)); $mainForm.Controls.Add($loginPanel)
$mainForm.Add_Resize({ if ($loginPanel.Visible) { $loginPanel.Left = ($mainForm.ClientSize.Width - $loginPanel.Width) / 2; $loginPanel.Top = ($mainForm.ClientSize.Height - $loginPanel.Height) / 3 } })

# --- Status Bar ---
$statusBar = New-Object System.Windows.Forms.StatusBar; $statusBar.Text = "Ready. Please login."; $mainForm.Controls.Add($statusBar)

# --- Main Tab Control ---
$tabControlMain = New-Object System.Windows.Forms.TabControl; $tabControlMain.Location = New-Object System.Drawing.Point(10, 10); $tabControlMain.Size = New-Object System.Drawing.Size(($mainForm.ClientSize.Width - 20), ($mainForm.ClientSize.Height - $statusBar.Height - 25)); $tabControlMain.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right; $tabControlMain.Visible = $false; $tabControlMain.Padding = New-Object System.Drawing.Point(10, 5); $mainForm.Controls.Add($tabControlMain)

# --- Tab Pages ---
$tabPageQuote = New-Object System.Windows.Forms.TabPage; $tabPageQuote.Text = "Quote"; $tabPageQuote.BackColor = $colorPanel; $tabPageQuote.Padding = $paddingMedium
$tabPageSettings = New-Object System.Windows.Forms.TabPage; $tabPageSettings.Text = "Settings"; $tabPageSettings.BackColor = $colorPanel; $tabPageSettings.Padding = $paddingMedium
$tabPageReports = New-Object System.Windows.Forms.TabPage; $tabPageReports.Text = "Reports"; $tabPageReports.BackColor = $colorPanel; $tabPageReports.Padding = $paddingMedium
$tabControlMain.Controls.AddRange(@($tabPageQuote, $tabPageSettings, $tabPageReports))

# ============================================================
# Quote Tab UI (Multi-Item Version)
# ============================================================
$singleQuotePanel = New-Object System.Windows.Forms.Panel; $singleQuotePanel.Dock = [System.Windows.Forms.DockStyle]::Fill; $singleQuotePanel.BackColor = $colorPanel; $tabPageQuote.Controls.Add($singleQuotePanel)

# --- Customer Selection (Top Right) ---
$labelSelectCustomer_Quote = New-Object System.Windows.Forms.Label; $labelSelectCustomer_Quote.Text = "Select Customer:"; $labelSelectCustomer_Quote.Location = New-Object System.Drawing.Point(550, 15); $labelSelectCustomer_Quote.AutoSize = $true; $labelSelectCustomer_Quote.Font = $fontBold
$comboBoxSelectCustomer_Quote = New-Object System.Windows.Forms.ComboBox; $comboBoxSelectCustomer_Quote.Location = New-Object System.Drawing.Point(550, 35); $comboBoxSelectCustomer_Quote.Size = New-Object System.Drawing.Size(200, 23); $comboBoxSelectCustomer_Quote.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; $comboBoxSelectCustomer_Quote.Enabled = $false;

# --- Input Fields Layout ---
[int]$col1X = 15; [int]$col2X = 300; [int]$labelWidth = 90; [int]$textBoxWidth = 160; [int]$rowHeight = 30
$currentRowY = 15

# Origin/Destination Headers
$labelOriginHeader = New-Object System.Windows.Forms.Label; $labelOriginHeader.Text = "Origin Details:"; $labelOriginHeader.Location = New-Object System.Drawing.Point($col1X, $currentRowY); $labelOriginHeader.Font = $fontBold; $labelOriginHeader.AutoSize = $true; $labelOriginHeader.ForeColor = $colorPrimary
$labelDestHeader = New-Object System.Windows.Forms.Label; $labelDestHeader.Text = "Destination Details:"; $labelDestHeader.Location = New-Object System.Drawing.Point($col2X, $currentRowY); $labelDestHeader.Font = $fontBold; $labelDestHeader.AutoSize = $true; $labelDestHeader.ForeColor = $colorPrimary
$currentRowY += $rowHeight

# ZIP Code Row
$labelOriginZip = New-Object System.Windows.Forms.Label; $labelOriginZip.Text = "ZIP Code:"; $labelOriginZip.Location = New-Object System.Drawing.Point($col1X, $currentRowY); $labelOriginZip.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelOriginZip.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $textboxOriginZip = New-Object System.Windows.Forms.TextBox; $textboxOriginZip.Location = New-Object System.Drawing.Point(($col1X + $labelWidth + 5), $currentRowY); $textboxOriginZip.Size = New-Object System.Drawing.Size($textBoxWidth, 23); $textboxOriginZip.MaxLength = 6; $textboxOriginZip.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle 
$labelDestZip = New-Object System.Windows.Forms.Label; $labelDestZip.Text = "ZIP Code:"; $labelDestZip.Location = New-Object System.Drawing.Point($col2X, $currentRowY); $labelDestZip.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelDestZip.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $textboxDestZip = New-Object System.Windows.Forms.TextBox; $textboxDestZip.Location = New-Object System.Drawing.Point(($col2X + $labelWidth + 5), $currentRowY); $textboxDestZip.Size = New-Object System.Drawing.Size($textBoxWidth, 23); $textboxDestZip.MaxLength = 6; $textboxDestZip.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle; $currentRowY += $rowHeight

# City Row
$labelOriginCity = New-Object System.Windows.Forms.Label; $labelOriginCity.Text = "City:"; $labelOriginCity.Location = New-Object System.Drawing.Point($col1X, $currentRowY); $labelOriginCity.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelOriginCity.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $textboxOriginCity = New-Object System.Windows.Forms.TextBox; $textboxOriginCity.Location = New-Object System.Drawing.Point(($col1X + $labelWidth + 5), $currentRowY); $textboxOriginCity.Size = New-Object System.Drawing.Size($textBoxWidth, 23); $textboxOriginCity.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelDestCity = New-Object System.Windows.Forms.Label; $labelDestCity.Text = "City:"; $labelDestCity.Location = New-Object System.Drawing.Point($col2X, $currentRowY); $labelDestCity.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelDestCity.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $textboxDestCity = New-Object System.Windows.Forms.TextBox; $textboxDestCity.Location = New-Object System.Drawing.Point(($col2X + $labelWidth + 5), $currentRowY); $textboxDestCity.Size = New-Object System.Drawing.Size($textBoxWidth, 23); $textboxDestCity.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle; $currentRowY += $rowHeight

# State Row
$labelOriginState = New-Object System.Windows.Forms.Label; $labelOriginState.Text = "State (2 Ltr):"; $labelOriginState.Location = New-Object System.Drawing.Point($col1X, $currentRowY); $labelOriginState.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelOriginState.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $textboxOriginState = New-Object System.Windows.Forms.TextBox; $textboxOriginState.Location = New-Object System.Drawing.Point(($col1X + $labelWidth + 5), $currentRowY); $textboxOriginState.Size = New-Object System.Drawing.Size(50, 23); $textboxOriginState.MaxLength = 2; $textboxOriginState.CharacterCasing = [System.Windows.Forms.CharacterCasing]::Upper; $textboxOriginState.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelDestState = New-Object System.Windows.Forms.Label; $labelDestState.Text = "State (2 Ltr):"; $labelDestState.Location = New-Object System.Drawing.Point($col2X, $currentRowY); $labelDestState.Size = New-Object System.Drawing.Size($labelWidth, 23); $labelDestState.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight; $textboxDestState = New-Object System.Windows.Forms.TextBox; $textboxDestState.Location = New-Object System.Drawing.Point(($col2X + $labelWidth + 5), $currentRowY); $textboxDestState.Size = New-Object System.Drawing.Size(50, 23); $textboxDestState.MaxLength = 2; $textboxDestState.CharacterCasing = [System.Windows.Forms.CharacterCasing]::Upper; $textboxDestState.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle; $currentRowY += $rowHeight + 10

# --- Commodity Grid ---
$labelCommodities = New-Object System.Windows.Forms.Label; $labelCommodities.Text = "Commodities:"; $labelCommodities.Location = New-Object System.Drawing.Point($col1X, $currentRowY); $labelCommodities.Font = $fontBold; $labelCommodities.AutoSize = $true; $labelCommodities.ForeColor = $colorPrimary
$currentRowY += 25

$dataGridViewCommodities = New-Object System.Windows.Forms.DataGridView
$dataGridViewCommodities.Location = New-Object System.Drawing.Point($col1X, $currentRowY)
$dataGridViewCommodities.Size = New-Object System.Drawing.Size(720, 150) 
$dataGridViewCommodities.ColumnHeadersHeightSizeMode = [System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::AutoSize
$dataGridViewCommodities.AllowUserToAddRows = $false 
$dataGridViewCommodities.AllowUserToDeleteRows = $false 
$dataGridViewCommodities.EditMode = [System.Windows.Forms.DataGridViewEditMode]::EditOnEnter
$dataGridViewCommodities.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dataGridViewCommodities.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right

# Define Columns
$colPieces = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colPieces.Name = "Pieces"; $colPieces.HeaderText = "Pcs"; $colPieces.Width = 40
$colClass = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colClass.Name = "Class"; $colClass.HeaderText = "Class"; $colClass.Width = 50
$colWeight = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colWeight.Name = "Weight"; $colWeight.HeaderText = "Weight (lbs)"; $colWeight.Width = 80
$colLength = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colLength.Name = "Length"; $colLength.HeaderText = "L (in)"; $colLength.Width = 50
$colWidth = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colWidth.Name = "Width"; $colWidth.HeaderText = "W (in)"; $colWidth.Width = 50
$colHeight = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colHeight.Name = "Height"; $colHeight.HeaderText = "H (in)"; $colHeight.Width = 50
$colDesc = New-Object System.Windows.Forms.DataGridViewTextBoxColumn; $colDesc.Name = "Description"; $colDesc.HeaderText = "Description"; $colDesc.AutoSizeMode = [System.Windows.Forms.DataGridViewAutoSizeColumnMode]::Fill
$colStack = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn; $colStack.Name = "Stackable"; $colStack.HeaderText = "Stack"; $colStack.Width = 50; $colStack.TrueValue = 'Y'; $colStack.FalseValue = 'N'; $colStack.IndeterminateValue = 'N' 

$dataGridViewCommodities.Columns.AddRange($colPieces, $colClass, $colWeight, $colLength, $colWidth, $colHeight, $colDesc, $colStack)
$currentRowY += $dataGridViewCommodities.Height + 5

# Add/Remove Buttons for Grid
$buttonAddItem = New-Object System.Windows.Forms.Button; $buttonAddItem.Text = "Add Item"; $buttonAddItem.Location = New-Object System.Drawing.Point($col1X, $currentRowY); $buttonAddItem.Size = New-Object System.Drawing.Size(90, 28); $buttonAddItem.Font = $fontRegular; $buttonAddItem.BackColor = $colorButtonBack; $buttonAddItem.ForeColor = $colorButtonFore; $buttonAddItem.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $buttonAddItem.FlatAppearance.BorderSize = 1; $buttonAddItem.FlatAppearance.BorderColor = $colorButtonBorder
$buttonRemoveItem = New-Object System.Windows.Forms.Button; $buttonRemoveItem.Text = "Remove Item"; $buttonRemoveItem.Location = New-Object System.Drawing.Point(($col1X + $buttonAddItem.Width + 10), $currentRowY); $buttonRemoveItem.Size = New-Object System.Drawing.Size(110, 28); $buttonRemoveItem.Font = $fontRegular; $buttonRemoveItem.BackColor = $colorButtonBack; $buttonRemoveItem.ForeColor = $colorButtonFore; $buttonRemoveItem.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $buttonRemoveItem.FlatAppearance.BorderSize = 1; $buttonRemoveItem.FlatAppearance.BorderColor = $colorButtonBorder
$currentRowY += $buttonAddItem.Height + 15

# Get Quote Button
$buttonGetQuote = New-Object System.Windows.Forms.Button; $buttonGetQuote.Text = "Get Quote"; $buttonGetQuote.Size = New-Object System.Drawing.Size(110, 35); $buttonGetQuote.Location = New-Object System.Drawing.Point(($singleQuotePanel.ClientSize.Width - $buttonGetQuote.Width - 20), ($currentRowY - $buttonGetQuote.Height - 10)); $buttonGetQuote.Font = $fontBold; $buttonGetQuote.BackColor = $colorButtonBack; $buttonGetQuote.ForeColor = $colorButtonFore; $buttonGetQuote.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $buttonGetQuote.FlatAppearance.BorderSize = 1; $buttonGetQuote.FlatAppearance.BorderColor = $colorButtonBorder; $buttonGetQuote.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right

# Results Area
$labelResults = New-Object System.Windows.Forms.Label; $labelResults.Text = "Quote Results:"; $labelResults.Location = New-Object System.Drawing.Point($col1X, $currentRowY); $labelResults.Font = $fontBold; $labelResults.AutoSize = $true; $labelResults.ForeColor = $colorPrimary
$currentRowY += 25
$textboxResults = New-Object System.Windows.Forms.TextBox; $textboxResults.Location = New-Object System.Drawing.Point($col1X, $currentRowY); $textboxResults.Size = New-Object System.Drawing.Size(($singleQuotePanel.ClientSize.Width - $col1X - 15), ($singleQuotePanel.ClientSize.Height - $currentRowY - 15)); $textboxResults.Multiline = $true; $textboxResults.ReadOnly = $true; $textboxResults.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical; $textboxResults.Font = $fontMono; $textboxResults.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right; $textboxResults.BackColor = $colorPanel; $textboxResults.ForeColor = $colorText

# Add Controls to Panel
$singleQuotePanel.Controls.AddRange(@(
    $labelSelectCustomer_Quote, $comboBoxSelectCustomer_Quote,
    $labelOriginHeader, $labelDestHeader,
    $labelOriginZip, $textboxOriginZip, $labelDestZip, $textboxDestZip,
    $labelOriginCity, $textboxOriginCity, $labelDestCity, $textboxDestCity,
    $labelOriginState, $textboxOriginState, $labelDestState, $textboxDestState,
    $labelCommodities, $dataGridViewCommodities, $buttonAddItem, $buttonRemoveItem,
    $buttonGetQuote,
    $labelResults, $textboxResults
))

# ============================================================
# Settings Tab UI
# ============================================================
$settingsPanel = New-Object System.Windows.Forms.Panel; $settingsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill; $settingsPanel.BackColor = $colorPanel; $tabPageSettings.Controls.Add($settingsPanel)
$labelSelectCustomer_Settings = New-Object System.Windows.Forms.Label; $labelSelectCustomer_Settings.Text = "Select Customer:"; $labelSelectCustomer_Settings.Location = New-Object System.Drawing.Point(460, 15); $labelSelectCustomer_Settings.AutoSize = $true; $labelSelectCustomer_Settings.Font = $fontBold; $labelSelectCustomer_Settings.ForeColor = $colorText 
$comboBoxSelectCustomer_Settings = New-Object System.Windows.Forms.ComboBox; $comboBoxSelectCustomer_Settings.Location = New-Object System.Drawing.Point(460, 35); $comboBoxSelectCustomer_Settings.Size = New-Object System.Drawing.Size(250, 23); $comboBoxSelectCustomer_Settings.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; $comboBoxSelectCustomer_Settings.Enabled = $false; $comboBoxSelectCustomer_Settings.Font = $fontRegular
$groupBoxCarrierSelect = New-Object System.Windows.Forms.GroupBox; $groupBoxCarrierSelect.Text = "Select Carrier"; $groupBoxCarrierSelect.Location = New-Object System.Drawing.Point(15, 15); $groupBoxCarrierSelect.Size = New-Object System.Drawing.Size(430, 55); $groupBoxCarrierSelect.Font = $fontRegular; $groupBoxCarrierSelect.ForeColor = $colorText 
$radioCentral = New-Object System.Windows.Forms.RadioButton; $radioCentral.Text = "Central"; $radioCentral.Location = New-Object System.Drawing.Point(15, 22); $radioCentral.AutoSize = $true; $radioCentral.Checked = $true; $radioCentral.Font = $fontRegular; $radioCentral.ForeColor = $colorText
$radioSAIA = New-Object System.Windows.Forms.RadioButton; $radioSAIA.Text = "SAIA"; $radioSAIA.Location = New-Object System.Drawing.Point(90, 22); $radioSAIA.AutoSize = $true; $radioSAIA.Font = $fontRegular; $radioSAIA.ForeColor = $colorText 
$radioRL = New-Object System.Windows.Forms.RadioButton; $radioRL.Text = "R+L"; $radioRL.Location = New-Object System.Drawing.Point(160, 22); $radioRL.AutoSize = $true; $radioRL.Font = $fontRegular; $radioRL.ForeColor = $colorText 
$radioAveritt = New-Object System.Windows.Forms.RadioButton; $radioAveritt.Text = "Averitt"; $radioAveritt.Location = New-Object System.Drawing.Point(230, 22); $radioAveritt.AutoSize = $true; $radioAveritt.Font = $fontRegular; $radioAveritt.ForeColor = $colorText 
$radioAAACooper = New-Object System.Windows.Forms.RadioButton; $radioAAACooper.Text = "AAA Cooper"; $radioAAACooper.Location = New-Object System.Drawing.Point(310, 22); $radioAAACooper.AutoSize = $true; $radioAAACooper.Font = $fontRegular; $radioAAACooper.ForeColor = $colorText 
$groupBoxCarrierSelect.Controls.AddRange(@($radioCentral, $radioSAIA, $radioRL, $radioAveritt, $radioAAACooper))
$labelTariffList = New-Object System.Windows.Forms.Label; $labelTariffList.Text = "Permitted Tariffs && Margins (for Selected Customer):"; $labelTariffList.Location = New-Object System.Drawing.Point(15, 80); $labelTariffList.AutoSize = $true; $labelTariffList.Font = $fontBold; $labelTariffList.ForeColor = $colorPrimary
$listBoxTariffs = New-Object System.Windows.Forms.ListBox; $listBoxTariffs.Location = New-Object System.Drawing.Point(15, 105); $listBoxTariffs.Size = New-Object System.Drawing.Size(350, 220); $listBoxTariffs.Font = $fontMono; $listBoxTariffs.IntegralHeight = $false; $listBoxTariffs.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$groupBoxSetMargin = New-Object System.Windows.Forms.GroupBox; $groupBoxSetMargin.Text = "Set Margin for Selected Tariff"; $groupBoxSetMargin.Location = New-Object System.Drawing.Point(380, 100); $groupBoxSetMargin.Size = New-Object System.Drawing.Size(330, 130); $groupBoxSetMargin.Font = $fontRegular; $groupBoxSetMargin.ForeColor = $colorText 
$labelSelectedTariff = New-Object System.Windows.Forms.Label; $labelSelectedTariff.Text = "Selected: (None)"; $labelSelectedTariff.Location = New-Object System.Drawing.Point(15, 28); $labelSelectedTariff.AutoSize = $true; $labelSelectedTariff.MaximumSize = New-Object System.Drawing.Size(($groupBoxSetMargin.Width - 30), 0); $labelSelectedTariff.Font = $fontBold; $labelSelectedTariff.ForeColor = $colorText
$labelNewMargin = New-Object System.Windows.Forms.Label; $labelNewMargin.Text = "New Margin %:"; $labelNewMargin.Location = New-Object System.Drawing.Point(15, 60); $labelNewMargin.AutoSize = $true; $labelNewMargin.ForeColor = $colorText
$textBoxNewMargin = New-Object System.Windows.Forms.TextBox; $textBoxNewMargin.Location = New-Object System.Drawing.Point(120, 57); $textBoxNewMargin.Size = New-Object System.Drawing.Size(70, 23); $textBoxNewMargin.Enabled = $false; $textBoxNewMargin.Font = $fontRegular; $textBoxNewMargin.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$buttonSetMargin = New-Object System.Windows.Forms.Button; $buttonSetMargin.Text = "Set Margin"; $buttonSetMargin.Location = New-Object System.Drawing.Point(75, 90); $buttonSetMargin.Size = New-Object System.Drawing.Size(100, 30); $buttonSetMargin.Enabled = $false; $buttonSetMargin.Font = $fontBold; $buttonSetMargin.BackColor = $colorButtonBack; $buttonSetMargin.ForeColor = $colorButtonFore; $buttonSetMargin.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $buttonSetMargin.FlatAppearance.BorderSize = 1; $buttonSetMargin.FlatAppearance.BorderColor = $colorButtonBorder
$groupBoxSetMargin.Controls.AddRange(@($labelSelectedTariff, $labelNewMargin, $textBoxNewMargin, $buttonSetMargin))
$settingsPanel.Controls.AddRange(@( $labelSelectCustomer_Settings, $comboBoxSelectCustomer_Settings, $groupBoxCarrierSelect, $labelTariffList, $listBoxTariffs, $groupBoxSetMargin ))

# ============================================================
# Reports Tab UI
# ============================================================
$reportsPanel = New-Object System.Windows.Forms.Panel; $reportsPanel.Dock = [System.Windows.Forms.DockStyle]::Fill; $reportsPanel.BackColor = $colorPanel; $tabPageReports.Controls.Add($reportsPanel)
[int]$reportsPanelX = 15; [int]$reportsPanelY = 15; [int]$reportsControlSpacing = 35; [int]$reportsLabelWidth = 150; [int]$reportsInputWidth = 250; [int]$reportsListBoxWidth = 250; [int]$reportsListBoxHeight = 100; [int]$groupBoxAspInputHeight = 55; [int]$checkBoxApplyMarginsHeight = 25; [int]$verticalPaddingBetweenControls = 10; $currentY = $reportsPanelY
$labelSelectCustomer_Reports = New-Object System.Windows.Forms.Label; $labelSelectCustomer_Reports.Text = "Select Customer:"; $labelSelectCustomer_Reports.Location = New-Object System.Drawing.Point($reportsPanelX, $currentY); $labelSelectCustomer_Reports.AutoSize = $true; $labelSelectCustomer_Reports.Font = $fontBold; $labelSelectCustomer_Reports.ForeColor = $colorText
$comboBoxSelectCustomer_Reports = New-Object System.Windows.Forms.ComboBox; $comboBoxSelectCustomer_Reports.Location = New-Object System.Drawing.Point(($reportsPanelX + $reportsLabelWidth + 5), ($currentY - 3)); $comboBoxSelectCustomer_Reports.Size = New-Object System.Drawing.Size($reportsInputWidth, 23); $comboBoxSelectCustomer_Reports.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; $comboBoxSelectCustomer_Reports.Enabled = $false; $comboBoxSelectCustomer_Reports.Font = $fontRegular; $currentY += $reportsControlSpacing
$labelReportType = New-Object System.Windows.Forms.Label; $labelReportType.Text = "Select Report Type:"; $labelReportType.Location = New-Object System.Drawing.Point($reportsPanelX, $currentY); $labelReportType.AutoSize = $true; $labelReportType.Font = $fontBold; $labelReportType.ForeColor = $colorText
$comboBoxReportType = New-Object System.Windows.Forms.ComboBox; $comboBoxReportType.Location = New-Object System.Drawing.Point(($reportsPanelX + $reportsLabelWidth + 5), ($currentY - 3)); $comboBoxReportType.Size = New-Object System.Drawing.Size($reportsInputWidth, 23); $comboBoxReportType.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList; $comboBoxReportType.Enabled = $false; $comboBoxReportType.Font = $fontRegular; $comboBoxReportType.Items.AddRange(@( "Carrier Comparison", "Avg Required Margin", "Required Margin for ASP", "Cross-Carrier ASP Analysis", "Margins by History" )); $currentY += $reportsControlSpacing
$groupBoxReportCarrierSelect = New-Object System.Windows.Forms.GroupBox; $groupBoxReportCarrierSelect.Text = "Select Carrier"; $groupBoxReportCarrierSelect.Size = New-Object System.Drawing.Size(430, 55); $groupBoxReportCarrierSelect.Font = $fontRegular; $groupBoxReportCarrierSelect.ForeColor = $colorText; $groupBoxReportCarrierSelect.Visible = $false 
$radioReportCentral = New-Object System.Windows.Forms.RadioButton; $radioReportCentral.Text = "Central"; $radioReportCentral.Location = New-Object System.Drawing.Point(15, 22); $radioReportCentral.AutoSize = $true; $radioReportCentral.Checked = $true; $radioReportCentral.Font = $fontRegular; $radioReportCentral.ForeColor = $colorText
$radioReportSAIA = New-Object System.Windows.Forms.RadioButton; $radioReportSAIA.Text = "SAIA"; $radioReportSAIA.Location = New-Object System.Drawing.Point(90, 22); $radioReportSAIA.AutoSize = $true; $radioReportSAIA.Font = $fontRegular; $radioReportSAIA.ForeColor = $colorText 
$radioReportRL = New-Object System.Windows.Forms.RadioButton; $radioReportRL.Text = "R+L"; $radioReportRL.Location = New-Object System.Drawing.Point(160, 22); $radioReportRL.AutoSize = $true; $radioReportRL.Font = $fontRegular; $radioReportRL.ForeColor = $colorText 
$radioReportAveritt = New-Object System.Windows.Forms.RadioButton; $radioReportAveritt.Text = "Averitt"; $radioReportAveritt.Location = New-Object System.Drawing.Point(230, 22); $radioReportAveritt.AutoSize = $true; $radioReportAveritt.Font = $fontRegular; $radioReportAveritt.ForeColor = $colorText 
$radioReportAAACooper = New-Object System.Windows.Forms.RadioButton; $radioReportAAACooper.Text = "AAA Cooper"; $radioReportAAACooper.Location = New-Object System.Drawing.Point(310, 22); $radioReportAAACooper.AutoSize = $true; $radioReportAAACooper.Font = $fontRegular; $radioReportAAACooper.ForeColor = $colorText 
$groupBoxReportCarrierSelect.Controls.AddRange(@($radioReportCentral, $radioReportSAIA, $radioReportRL, $radioReportAveritt, $radioReportAAACooper))
$groupBoxReportTariffSelect = New-Object System.Windows.Forms.GroupBox; $groupBoxReportTariffSelect.Text = "Select Tariff(s)"; $groupBoxReportTariffSelect.Size = New-Object System.Drawing.Size(550, ($reportsListBoxHeight + 40)); $groupBoxReportTariffSelect.Font = $fontRegular; $groupBoxReportTariffSelect.ForeColor = $colorText; $groupBoxReportTariffSelect.Visible = $false
$labelReportTariff1 = New-Object System.Windows.Forms.Label; $labelReportTariff1.Text = "Tariff 1 (Base/Cost):"; $labelReportTariff1.Location = New-Object System.Drawing.Point(10, 25); $labelReportTariff1.AutoSize = $true; $labelReportTariff1.Font = $fontBold; $labelReportTariff1.ForeColor = $colorText
$listBoxReportTariff1 = New-Object System.Windows.Forms.ListBox; $listBoxReportTariff1.Location = New-Object System.Drawing.Point(10, 45); $listBoxReportTariff1.Size = New-Object System.Drawing.Size($reportsListBoxWidth, $reportsListBoxHeight); $listBoxReportTariff1.Font = $fontMono; $listBoxReportTariff1.IntegralHeight = $false; $listBoxReportTariff1.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$labelReportTariff2 = New-Object System.Windows.Forms.Label; $labelReportTariff2.Text = "Tariff 2 (Comparison):"; $labelReportTariff2.Location = New-Object System.Drawing.Point(280, 25); $labelReportTariff2.AutoSize = $true; $labelReportTariff2.Font = $fontBold; $labelReportTariff2.ForeColor = $colorText; $labelReportTariff2.Visible = $false
$listBoxReportTariff2 = New-Object System.Windows.Forms.ListBox; $listBoxReportTariff2.Location = New-Object System.Drawing.Point(280, 45); $listBoxReportTariff2.Size = New-Object System.Drawing.Size($reportsListBoxWidth, $reportsListBoxHeight); $listBoxReportTariff2.Font = $fontMono; $listBoxReportTariff2.IntegralHeight = $false; $listBoxReportTariff2.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle; $listBoxReportTariff2.Visible = $false
$groupBoxReportTariffSelect.Controls.AddRange(@($labelReportTariff1, $listBoxReportTariff1, $labelReportTariff2, $listBoxReportTariff2))
$labelSelectCsv = New-Object System.Windows.Forms.Label; $labelSelectCsv.Text = "Select Data File (CSV):"; $labelSelectCsv.AutoSize = $true; $labelSelectCsv.Font = $fontBold; $labelSelectCsv.ForeColor = $colorText; $labelSelectCsv.Visible = $false
$textboxCsvPath = New-Object System.Windows.Forms.TextBox; $textboxCsvPath.Size = New-Object System.Drawing.Size($reportsInputWidth, 23); $textboxCsvPath.ReadOnly = $true; $textboxCsvPath.Font = $fontRegular; $textboxCsvPath.BackColor = $colorPanel; $textboxCsvPath.Visible = $false
$buttonSelectCsv = New-Object System.Windows.Forms.Button; $buttonSelectCsv.Text = "Browse..."; $buttonSelectCsv.Size = New-Object System.Drawing.Size(80, 28); $buttonSelectCsv.Font = $fontRegular; $buttonSelectCsv.BackColor = $colorButtonBack; $buttonSelectCsv.ForeColor = $colorButtonFore; $buttonSelectCsv.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $buttonSelectCsv.FlatAppearance.BorderSize = 1; $buttonSelectCsv.FlatAppearance.BorderColor = $colorButtonBorder; $buttonSelectCsv.Enabled = $false; $buttonSelectCsv.Visible = $false
$groupBoxReportAspInput = New-Object System.Windows.Forms.GroupBox; $groupBoxReportAspInput.Text = "Desired Selling Price"; $groupBoxReportAspInput.Size = New-Object System.Drawing.Size(410, $groupBoxAspInputHeight); $groupBoxReportAspInput.Font = $fontRegular; $groupBoxReportAspInput.ForeColor = $colorText; $groupBoxReportAspInput.Visible = $false
$labelDesiredAsp = New-Object System.Windows.Forms.Label; $labelDesiredAsp.Text = "Desired Avg Selling Price ($):"; $labelDesiredAsp.Location = New-Object System.Drawing.Point(10, 25); $labelDesiredAsp.AutoSize = $true; $labelDesiredAsp.ForeColor = $colorText
$textBoxDesiredAsp = New-Object System.Windows.Forms.TextBox; $textBoxDesiredAsp.Location = New-Object System.Drawing.Point(190, 22); $textBoxDesiredAsp.Size = New-Object System.Drawing.Size(100, 23); $textBoxDesiredAsp.Font = $fontRegular; $textBoxDesiredAsp.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$groupBoxReportAspInput.Controls.AddRange(@($labelDesiredAsp, $textBoxDesiredAsp))
$checkBoxApplyMargins = New-Object System.Windows.Forms.CheckBox; $checkBoxApplyMargins.Text = "Apply Calculated Margins to Tariff Files (Use with Caution!)"; $checkBoxApplyMargins.AutoSize = $true; $checkBoxApplyMargins.Font = $fontRegular; $checkBoxApplyMargins.ForeColor = [System.Drawing.Color]::DarkRed; $checkBoxApplyMargins.Visible = $false
$buttonRunReport = New-Object System.Windows.Forms.Button; $buttonRunReport.Text = "Run Report"; $buttonRunReport.Size = New-Object System.Drawing.Size(120, 35); $buttonRunReport.Font = $fontBold; $buttonRunReport.BackColor = $colorButtonBack; $buttonRunReport.ForeColor = $colorButtonFore; $buttonRunReport.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat; $buttonRunReport.FlatAppearance.BorderSize = 1; $buttonRunReport.FlatAppearance.BorderColor = $colorButtonBorder; $buttonRunReport.Enabled = $false
$textboxReportResults = New-Object System.Windows.Forms.TextBox; $textboxReportResults.Multiline = $true; $textboxReportResults.ReadOnly = $true; $textboxReportResults.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical; $textboxReportResults.Font = $fontMono; $textboxReportResults.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right; $textboxReportResults.BackColor = $colorPanel; $textboxReportResults.ForeColor = $colorText
$reportsPanel.Controls.AddRange(@( $labelSelectCustomer_Reports, $comboBoxSelectCustomer_Reports, $labelReportType, $comboBoxReportType, $groupBoxReportCarrierSelect, $groupBoxReportTariffSelect, $labelSelectCsv, $textboxCsvPath, $buttonSelectCsv, $groupBoxReportAspInput, $checkBoxApplyMargins, $buttonRunReport, $textboxReportResults ))

# ============================================================
# Event Handlers Section
# ============================================================

$buttonLogin.Add_Click({
    param($sender, $e)
    $username = $textboxUsername.Text; $password = $textboxPassword.Text
    if (-not (Get-Command Test-PasswordHash -EA SilentlyContinue)) { [System.Windows.Forms.MessageBox]::Show("FATAL: Test-PasswordHash missing."); return }
    $script:currentUserProfile = $null
    try { $script:currentUserProfile = Authenticate-User -Username $username -PasswordPlainText $password -UserAccountsFolderPath $script:userAccountsFolderPath } catch { [System.Windows.Forms.MessageBox]::Show("Auth Error: $($_.Exception.Message)") }

    if ($script:currentUserProfile) {
        $statusBar.Text = "Welcome, $($script:currentUserProfile.Username)."
        $loginPanel.Visible = $false; $tabControlMain.Visible = $true
        $comboBoxSelectCustomer_Quote.Items.Clear(); $comboBoxSelectCustomer_Settings.Items.Clear(); $comboBoxSelectCustomer_Reports.Items.Clear()
        $initialCustomerName = $null
        if ($script:allCustomerProfiles.Count -gt 0) {
            $customerNames = $script:allCustomerProfiles.Keys | Sort-Object
            $comboBoxSelectCustomer_Quote.Items.AddRange($customerNames); $comboBoxSelectCustomer_Settings.Items.AddRange($customerNames); $comboBoxSelectCustomer_Reports.Items.AddRange($customerNames)
            $comboBoxSelectCustomer_Quote.SelectedIndex = 0; $comboBoxSelectCustomer_Settings.SelectedIndex = 0; $comboBoxSelectCustomer_Reports.SelectedIndex = 0
            $comboBoxSelectCustomer_Quote.Enabled = $true; $comboBoxSelectCustomer_Settings.Enabled = $true; $comboBoxSelectCustomer_Reports.Enabled = $true
            $script:selectedCustomerProfile = $script:allCustomerProfiles[$customerNames[0]] 
            $initialCustomerName = $customerNames[0]
        } else {
            $noCustMsg = "No Customers Found"; $comboBoxSelectCustomer_Quote.Items.Add($noCustMsg); $comboBoxSelectCustomer_Settings.Items.Add($noCustMsg); $comboBoxSelectCustomer_Reports.Items.Add($noCustMsg)
            $comboBoxSelectCustomer_Quote.Enabled = $false; $comboBoxSelectCustomer_Settings.Enabled = $false; $comboBoxSelectCustomer_Reports.Enabled = $false
        }
        $script:currentUserReportsFolder = Join-Path $script:reportsBaseFolderPath $script:currentUserProfile.Username; Ensure-DirectoryExists $script:currentUserReportsFolder
        if ($null -ne $initialCustomerName) {
            $initialCarrier = "Central"; if ($radioSAIA.Checked){$initialCarrier="SAIA"} elseif ($radioRL.Checked){$initialCarrier="RL"} elseif ($radioAveritt.Checked){$initialCarrier="Averitt"} elseif ($radioAAACooper.Checked){$initialCarrier="AAACooper"}
            try { Populate-TariffListBox -SelectedCarrier $initialCarrier -ListBoxControl $listBoxTariffs -LabelControl $labelSelectedTariff -ButtonControl $buttonSetMargin -TextboxControl $textBoxNewMargin -SelectedCustomerName $initialCustomerName -AllCustomerProfiles $script:allCustomerProfiles } catch { $statusBar.Text = "Error init settings: $($_.Exception.Message)" }
        } else { $listBoxTariffs.Items.Clear(); $listBoxTariffs.Items.Add("No Customers") }
        $comboBoxReportType.Enabled = $true; $buttonSelectCsv.Enabled = $true; $buttonRunReport.Enabled = $true
        if ($comboBoxReportType.Items.Count -gt 0) { $comboBoxReportType.SelectedIndex = 0 } else { $comboBoxReportType_SelectedIndexChanged_ScriptBlock.Invoke($comboBoxReportType, [System.EventArgs]::Empty) } # Ensure UI updates if no items
        $textboxPassword.Clear()
    } else { if ($Error.Count -eq 0 -or ($Error[0].Exception.Message -notlike "*User account file not found*")) { $statusBar.Text = "Login failed."; [System.Windows.Forms.MessageBox]::Show("Login Failed.") }; $textboxPassword.Clear(); $textboxUsername.Focus() } # Avoid double message if auth func already warned
})

$customerChangedHandler = {
    param($sender, $e)
    $selectedName = $null; if ($sender.SelectedIndex -ge 0) { $selectedName = $sender.SelectedItem.ToString() }
    if ($null -ne $selectedName -and $script:allCustomerProfiles.ContainsKey($selectedName)) {
        $script:selectedCustomerProfile = $script:allCustomerProfiles[$selectedName] 
        $statusBar.Text = "Selected Customer: $selectedName"
        $comboBoxes = @($comboBoxSelectCustomer_Quote, $comboBoxSelectCustomer_Settings, $comboBoxSelectCustomer_Reports); foreach($cb in $comboBoxes){ if($cb -ne $sender -and $cb.SelectedItem -ne $selectedName){ try {$cb.SelectedItem = $selectedName} catch {} } }
        if ($tabControlMain.SelectedTab -eq $tabPageSettings) {
            $carrier="Central"; if ($radioSAIA.Checked){$carrier="SAIA"} elseif ($radioRL.Checked){$carrier="RL"} elseif ($radioAveritt.Checked){$carrier="Averitt"} elseif ($radioAAACooper.Checked){$carrier="AAACooper"}
            try { Populate-TariffListBox -SelectedCarrier $carrier -ListBoxControl $listBoxTariffs -LabelControl $labelSelectedTariff -ButtonControl $buttonSetMargin -TextboxControl $textBoxNewMargin -SelectedCustomerName $selectedName -AllCustomerProfiles $script:allCustomerProfiles } catch { $statusBar.Text = "Error refresh settings: $($_.Exception.Message)" }
        }
        if ($tabControlMain.SelectedTab -eq $tabPageReports) { try { $comboBoxReportType_SelectedIndexChanged_ScriptBlock.Invoke($comboBoxReportType, [System.EventArgs]::Empty) } catch { $statusBar.Text = "Error refresh reports: $($_.Exception.Message)" } }
    } else { $script:selectedCustomerProfile = $null; $statusBar.Text = "No valid customer selected."; $listBoxTariffs.Items.Clear(); $listBoxTariffs.Items.Add("Select Customer"); $listBoxReportTariff1.Items.Clear(); $listBoxReportTariff1.Items.Add("Select Customer"); $listBoxReportTariff2.Items.Clear(); $listBoxReportTariff2.Items.Add("Select Customer"); $labelSelectedTariff.Text = "Selected: (None)"; $buttonSetMargin.Enabled = $false; $textBoxNewMargin.Enabled = $false; $textBoxNewMargin.Clear() }
}
$comboBoxSelectCustomer_Quote.Add_SelectedIndexChanged($customerChangedHandler); $comboBoxSelectCustomer_Settings.Add_SelectedIndexChanged($customerChangedHandler); $comboBoxSelectCustomer_Reports.Add_SelectedIndexChanged($customerChangedHandler);

# --- Quote Tab Handlers ---
$buttonAddItem.Add_Click({
    $dataGridViewCommodities.Rows.Add("1", "70", "100", "48", "40", "36", "Item Description", 'Y') | Out-Null
})

$buttonRemoveItem.Add_Click({
    for ($i = $dataGridViewCommodities.SelectedRows.Count - 1; $i -ge 0; $i--) {
        $row = $dataGridViewCommodities.SelectedRows[$i]
        if (-not $row.IsNewRow) { 
            $dataGridViewCommodities.Rows.Remove($row)
        }
    }
})

$buttonGetQuote.Add_Click({
    param($sender, $e)
    if ($null -eq $script:selectedCustomerProfile) { [System.Windows.Forms.MessageBox]::Show("Please select a customer first.", "Customer Not Selected"); return };
    $textboxResults.Clear(); $textboxResults.Text = "Fetching rates for Customer: $($script:selectedCustomerProfile.CustomerName)..."; $mainForm.Refresh()

    $originZip = $textboxOriginZip.Text; $originCity = $textboxOriginCity.Text; $originState = $textboxOriginState.Text; $destZip = $textboxDestZip.Text; $destCity = $textboxDestCity.Text; $destState = $textboxDestState.Text;
    $validationErrors = [System.Collections.Generic.List[string]]::new();
    if (!($originZip -match '^\d{5,6}$')){$validationErrors.Add("Origin ZIP (5 or 6 digits).")}; if ([string]::IsNullOrWhiteSpace($originCity)){$validationErrors.Add("Origin City.")}; if (!($originState -match '^[A-Za-z]{2}$')){$validationErrors.Add("Origin State.")};
    if (!($destZip -match '^\d{5,6}$')){$validationErrors.Add("Dest ZIP (5 or 6 digits).")}; if ([string]::IsNullOrWhiteSpace($destCity)){$validationErrors.Add("Dest City.")}; if (!($destState -match '^[A-Za-z]{2}$')){$validationErrors.Add("Dest State.")};

    # --- Commodity Grid Validation and Data Collection (INTEGRATED SNIPPET) ---
    $commoditiesForApi = @() 
    $totalWeight = 0.0 
    $firstClass = $null 

    if ($dataGridViewCommodities.Rows.Count -eq 0) {
        $validationErrors.Add("At least one commodity item is required.")
    } else {
        foreach ($row in $dataGridViewCommodities.Rows) {
            if ($row.IsNewRow) { continue } 

            $rowNum = $row.Index + 1
            $pcs = $row.Cells["Pieces"].Value
            $cls = $row.Cells["Class"].Value
            $wt = $row.Cells["Weight"].Value
            $len = $row.Cells["Length"].Value
            $wid = $row.Cells["Width"].Value
            $hgt = $row.Cells["Height"].Value
            $desc = $row.Cells["Description"].Value
            $stack = $row.Cells["Stackable"].Value 

            $rowIsValid = $true
            if ($null -eq $pcs -or -not ($pcs -match '^\d+$') -or ([int]$pcs -le 0)) { $validationErrors.Add("Row ${rowNum}: Invalid Pieces."); $rowIsValid = $false }
            
            # CORRECTED/IMPROVED CLASS VALIDATION:
            if ([string]::IsNullOrWhiteSpace($cls)) {
                if (-not ([string]::IsNullOrWhiteSpace($wt)) -and ($wt -as [decimal]) -ne $null -and ([decimal]$wt -gt 0)) {
                    $validationErrors.Add("Row ${rowNum}: Class is required when Weight is provided."); $rowIsValid = $false 
                }
            } elseif (-not ($cls -match '^\d+(\.\d+)?$') -or ([double]$cls -lt 1) -or ([double]$cls -gt 500) ) { 
                $validationErrors.Add("Row ${rowNum}: Class must be a valid number (e.g., 50, 77.5, 100)."); $rowIsValid = $false 
            }

            if ($null -eq $wt -or -not ($wt -match '^\d+(\.\d+)?$') -or ([decimal]$wt -le 0)) { $validationErrors.Add("Row ${rowNum}: Invalid Weight."); $rowIsValid = $false }
            
            if (-not ([string]::IsNullOrWhiteSpace($len)) -and (-not ($len -match '^\d+(\.\d+)?$') -or ([decimal]$len -le 0))) { $validationErrors.Add("Row ${rowNum}: Invalid Length."); $rowIsValid = $false }
            if (-not ([string]::IsNullOrWhiteSpace($wid)) -and (-not ($wid -match '^\d+(\.\d+)?$') -or ([decimal]$wid -le 0))) { $validationErrors.Add("Row ${rowNum}: Invalid Width."); $rowIsValid = $false }
            if (-not ([string]::IsNullOrWhiteSpace($hgt)) -and (-not ($hgt -match '^\d+(\.\d+)?$') -or ([decimal]$hgt -le 0))) { $validationErrors.Add("Row ${rowNum}: Invalid Height."); $rowIsValid = $false }

            if ($null -eq $stack) { $stack = 'N' } 

            if ($rowIsValid) {
                 $commoditiesForApi += [PSCustomObject]@{
                    pieces         = [string]$pcs
                    classification = [string]$cls 
                    class          = [string]$cls 
                    itemClass      = [string]$cls 
                    weight         = [string]$wt
                    length         = if ([string]::IsNullOrWhiteSpace($len)) {$null} else {[string]$len} 
                    width          = if ([string]::IsNullOrWhiteSpace($wid)) {$null} else {[string]$wid}
                    height         = if ([string]::IsNullOrWhiteSpace($hgt)) {$null} else {[string]$hgt}
                    description    = if ([string]::IsNullOrWhiteSpace($desc)) { "Commodity" } else { [string]$desc }
                    packagingType  = "PAT" 
                    stackable      = if($stack -eq 'Y') {'Y'} else {'N'} 
                    NMFC = $null 
                    NMFCSub = $null 
                    HandlingUnits = [string]$pcs 
                    HandlingUnitType = if ($row.Cells["Pieces"].Value -gt 0 -and ($row.Cells["Description"].Value -match "pallet" -or "PAT" -match "pallet" )) {"Pallets"} else {"Other"} 
                    HazMat = $null 
                    CubeU = if ((-not [string]::IsNullOrWhiteSpace($len)) -and ($len -as [decimal]) -ne $null -and ([decimal]$len -gt 0) -and `
                                (-not [string]::IsNullOrWhiteSpace($wid)) -and ($wid -as [decimal]) -ne $null -and ([decimal]$wid -gt 0) -and `
                                (-not [string]::IsNullOrWhiteSpace($hgt)) -and ($hgt -as [decimal]) -ne $null -and ([decimal]$hgt -gt 0)) { "IN" } else { $null }
                }
                # Ensure $wt is valid decimal before adding to $totalWeight
                if ($wt -as [decimal] -ne $null) {
                    $totalWeight += [decimal]$wt
                }
                if ($null -eq $firstClass -and -not [string]::IsNullOrWhiteSpace($cls)) { $firstClass = [string]$cls }
            }
        } 
    } 
    # --- END OF INTEGRATED SNIPPET ---

    if ($validationErrors.Count -gt 0) { $errorMsg = "Invalid Input:`n$($validationErrors -join "`n")"; [System.Windows.Forms.MessageBox]::Show($errorMsg); $textboxResults.Text = $errorMsg; return };
    if ($commoditiesForApi.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show("No valid commodity items entered."); $textboxResults.Text = "No valid items."; return };

    $statusBar.Text = "Fetching rates...";
    $declaredValue=0.0; 
    $optionalShipmentDetails=[PSCustomObject]@{OriginCity=$originCity;OriginState=$originState;DestinationCity=$destCity;DestinationState=$destState; DeclaredValue=$declaredValue; OriginCountryCode="USA"; DestinationCountryCode="USA"}; 

    $resultsText=[System.Text.StringBuilder]::new(); $resultsText.AppendLine("="*20+" SHIPMENT QUOTE "+"="*20)|Out-Null; $resultsText.AppendLine("Quote Date: $(Get-Date)"); $resultsText.AppendLine("Broker User: $($script:currentUserProfile.Username)"); $resultsText.AppendLine("Customer:    $($script:selectedCustomerProfile.CustomerName)"); $resultsText.AppendLine("-"*58)|Out-Null; $resultsText.AppendLine("Origin:      $originCity, $originState $originZip"); $resultsText.AppendLine("Destination: $destCity, $destState $destZip"); $resultsText.AppendLine("Commodities: $($commoditiesForApi.Count) lines, Total Wt: $($totalWeight) lbs"); $resultsText.AppendLine("-"*58)|Out-Null; $resultsText.AppendLine("Carrier Options:")|Out-Null;

    $permittedCentralKeys=@{}; $permittedSAIAKeys=@{}; $permittedRLKeys=@{}; $permittedAverittKeys=@{}; $permittedAAACooperKeys=@{};
    try {
        $currentCustomerProfile = $script:selectedCustomerProfile; if(!$currentCustomerProfile){throw "Customer profile null."};
        $permittedCentralKeys=Get-PermittedKeys -AllKeys $script:allCentralKeys -AllowedKeyNames $currentCustomerProfile.AllowedCentralKeys
        $permittedSAIAKeys=Get-PermittedKeys -AllKeys $script:allSAIAKeys -AllowedKeyNames $currentCustomerProfile.AllowedSAIAKeys
        $permittedRLKeys=Get-PermittedKeys -AllKeys $script:allRLKeys -AllowedKeyNames $currentCustomerProfile.AllowedRLKeys
        $permittedAverittKeys=Get-PermittedKeys -AllKeys $script:allAverittKeys -AllowedKeyNames $currentCustomerProfile.AllowedAverittKeys
        $permittedAAACooperKeys=Get-PermittedKeys -AllKeys $script:allAAACooperKeys -AllowedKeyNames $currentCustomerProfile.AllowedAAACooperKeys
    } catch { [System.Windows.Forms.MessageBox]::Show("Error getting keys: $($_.Exception.Message)"); $textboxResults.Text = "Error getting keys."; return }

    $quoteTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'; $finalQuotes=@(); $CurrentVerbosePreference=$VerbosePreference; $VerbosePreference='SilentlyContinue';
    try {
        # Central
        if ($permittedCentralKeys.Count -gt 0) { $textboxResults.AppendText("Querying Central...`r`n"); $mainForm.Refresh(); $rates=@{}; foreach ($tariff in ($permittedCentralKeys.Keys|Sort)){ try {$keyData=$permittedCentralKeys[$tariff]; if(!$keyData.accessCode -or !$keyData.customerNumber){throw "cred missing."}; $centralShipmentData = [PSCustomObject]@{ "Origin Postal Code" = $originZip; "Destination Postal Code" = $destZip; "Commodities" = $commoditiesForApi; "Freight Class 1" = $firstClass}; $cost=Invoke-CentralTransportApi -ApiKeyData $keyData -OriginZip $originZip -DestinationZip $destZip -Commodities $commoditiesForApi; if($cost){$rates[$tariff]=$cost}}catch{$errMessage = if ($_.Exception) {$_.Exception.Message} else {$_}; $resultsText.AppendLine("  Central ($tariff): Err - $errMessage") }}; $lowest=Get-MinimumRate -RateResults $rates; if($lowest){ $data=$permittedCentralKeys[$lowest.TariffName]; $margin=$Global:DefaultMarginPercentage; if($data.MarginPercent){try{$margin=[double]$data.MarginPercent}catch{}}; $quote=Calculate-QuotePrice -LowestCarrierCost $lowest.Cost -OriginZip $originZip -DestinationZip $destZip -Weight $totalWeight -FreightClass $firstClass -MarginPercent $margin; if($quote.FinalPrice){$resultsText.AppendLine(("  {0,-20}: {1,-15} (Tariff: {2})" -f "Central", $quote.FinalPrice.ToString("C"), $lowest.TariffName)); $finalQuotes+=$quote; Write-QuoteToHistory -Carrier 'Central' -Tariff $lowest.TariffName -OriginZip $originZip -DestinationZip $destZip -Weight $totalWeight -FreightClass $firstClass -LowestCost $lowest.Cost -FinalQuotedPrice $quote.FinalPrice -QuoteTimestamp $quoteTimestamp -OriginZipFull $originZip -DestinationZipFull $destZip}else{$resultsText.AppendLine("  Central: Price Calc Err")}}else{$resultsText.AppendLine("  Central: No Rates")}}else{$resultsText.AppendLine("  Central: No Keys")}
        # SAIA
        if ($permittedSAIAKeys.Count -gt 0) { $textboxResults.AppendText("Querying SAIA...`r`n"); $mainForm.Refresh(); $rates=@{}; foreach ($tariff in ($permittedSAIAKeys.Keys|Sort)){ try {$keyData=$permittedSAIAKeys[$tariff]; $cost=Invoke-SAIAApi -KeyData $keyData -OriginZip $originZip -DestinationZip $destZip -OriginCity $originCity -OriginState $originState -DestinationCity $destCity -DestinationState $destState -Details $commoditiesForApi; if($cost){$rates[$tariff]=$cost}}catch{$errMessage = if ($_.Exception) {$_.Exception.Message} else {$_}; $resultsText.AppendLine("  SAIA ($tariff): Err - $errMessage") }}; $lowest=Get-MinimumRate -RateResults $rates; if($lowest){ $data=$permittedSAIAKeys[$lowest.TariffName]; $margin=$Global:DefaultMarginPercentage; if($data.MarginPercent){try{$margin=[double]$data.MarginPercent}catch{}}; $quote=Calculate-QuotePrice -LowestCarrierCost $lowest.Cost -OriginZip $originZip -DestinationZip $destZip -Weight $totalWeight -FreightClass $firstClass -MarginPercent $margin; if($quote.FinalPrice){$resultsText.AppendLine(("  {0,-20}: {1,-15} (Tariff: {2})" -f "SAIA", $quote.FinalPrice.ToString("C"), $lowest.TariffName)); $finalQuotes+=$quote; Write-QuoteToHistory -Carrier 'SAIA' -Tariff $lowest.TariffName -OriginZip $originZip -DestinationZip $destZip -Weight $totalWeight -FreightClass $firstClass -LowestCost $lowest.Cost -FinalQuotedPrice $quote.FinalPrice -QuoteTimestamp $quoteTimestamp -OriginZipFull $originZip -DestinationZipFull $destZip}else{$resultsText.AppendLine("  SAIA: Price Calc Err")}}else{$resultsText.AppendLine("  SAIA: No Rates")}}else{$resultsText.AppendLine("  SAIA: No Keys")}
        # R+L
        if ($permittedRLKeys.Count -gt 0) { $textboxResults.AppendText("Querying R+L...`r`n"); $mainForm.Refresh(); $rates=@{}; foreach ($tariff in ($permittedRLKeys.Keys|Sort)){ try {$keyData=$permittedRLKeys[$tariff]; if(!$keyData.APIKey){throw "cred missing."}; $cost=Invoke-RLApi -KeyData $keyData -OriginZip $originZip -DestinationZip $destZip -Commodities $commoditiesForApi -ShipmentDetails $optionalShipmentDetails; if($cost){$rates[$tariff]=$cost}}catch{$errMessage = if ($_.Exception) {$_.Exception.Message} else {$_}; $resultsText.AppendLine("  R+L ($tariff): Err - $errMessage") }}; $lowest=Get-MinimumRate -RateResults $rates; if($lowest){ $data=$permittedRLKeys[$lowest.TariffName]; $margin=$Global:DefaultMarginPercentage; if($data.MarginPercent){try{$margin=[double]$data.MarginPercent}catch{}}; $quote=Calculate-QuotePrice -LowestCarrierCost $lowest.Cost -OriginZip $originZip -DestinationZip $destZip -Weight $totalWeight -FreightClass $firstClass -MarginPercent $margin; if($quote.FinalPrice){$resultsText.AppendLine(("  {0,-20}: {1,-15} (Tariff: {2})" -f "R+L", $quote.FinalPrice.ToString("C"), $lowest.TariffName)); $finalQuotes+=$quote; Write-QuoteToHistory -Carrier 'R+L' -Tariff $lowest.TariffName -OriginZip $originZip -DestinationZip $destZip -Weight $totalWeight -FreightClass $firstClass -LowestCost $lowest.Cost -FinalQuotedPrice $quote.FinalPrice -QuoteTimestamp $quoteTimestamp -OriginZipFull $originZip -DestinationZipFull $destZip}else{$resultsText.AppendLine("  R+L: Price Calc Err")}}else{$resultsText.AppendLine("  R+L: No Rates")}}else{$resultsText.AppendLine("  R+L: No Keys")}
        # Averitt
        if ($permittedAverittKeys.Count -gt 0) {
            $textboxResults.AppendText("Querying Averitt...`r`n"); $mainForm.Refresh(); $rates=@{};
            foreach ($tariff in ($permittedAverittKeys.Keys|Sort)){
                try {
                    $keyData=$permittedAverittKeys[$tariff]; if(!$keyData.APIKey){throw "cred missing."};
                    $averittShipmentData=[PSCustomObject]@{
                        ServiceLevel = "STND"; PaymentTerms = "Prepaid"; PaymentPayer = "Shipper";
                        PickupDate = (Get-Date -Format 'yyyyMMdd');
                        OriginAccount = $keyData.AccountNumber; OriginCity = $originCity; OriginStateProvince = $originState; OriginPostalCode = $originZip; OriginCountry = "USA"; # Use AccountNumber from KeyData
                        DestinationAccount = $null; DestinationCity = $destCity; DestinationStateProvince = $destState; DestinationPostalCode = $destZip; DestinationCountry = "USA";
                        BillToAccount = $keyData.AccountNumber; BillToName = "Default BillTo"; BillToAddress = "123 Default"; BillToCity = "Default"; BillToStateProvince = "TN"; BillToPostalCode = "00000"; BillToCountry = "USA"; # Use AccountNumber from KeyData
                        Commodities = $commoditiesForApi; Accessorials = $null 
                    };
                    if(Get-Command Invoke-AverittApi -EA SilentlyContinue){$cost=Invoke-AverittApi -KeyData $keyData -ShipmentData $averittShipmentData; if($cost){$rates[$tariff]=$cost}} else { throw "Invoke-AverittApi missing."}
                } catch { $errMessage = if ($_.Exception) {$_.Exception.Message} else {$_}; $resultsText.AppendLine("  Averitt ($tariff): Err - $errMessage") }
            };
            $lowest=Get-MinimumRate -RateResults $rates;
            if($lowest){ $data=$permittedAverittKeys[$lowest.TariffName]; $margin=$Global:DefaultMarginPercentage; if($data.MarginPercent){try{$margin=[double]$data.MarginPercent}catch{}}; $quote=Calculate-QuotePrice -LowestCarrierCost $lowest.Cost -OriginZip $originZip -DestinationZip $destZip -Weight $totalWeight -FreightClass $firstClass -MarginPercent $margin; if($quote.FinalPrice){$resultsText.AppendLine(("  {0,-20}: {1,-15} (Tariff: {2})" -f "Averitt", $quote.FinalPrice.ToString("C"), $lowest.TariffName)); $finalQuotes+=$quote; Write-QuoteToHistory -Carrier 'Averitt' -Tariff $lowest.TariffName -OriginZip $originZip -DestinationZip $destZip -Weight $totalWeight -FreightClass $firstClass -LowestCost $lowest.Cost -FinalQuotedPrice $quote.FinalPrice -QuoteTimestamp $quoteTimestamp -OriginZipFull $originZip -DestinationZipFull $destZip}else{$resultsText.AppendLine("  Averitt: Price Calc Err")}}else{$resultsText.AppendLine("  Averitt: No Rates")}
        } else { $resultsText.AppendLine("  Averitt: No Keys") }

        # AAA Cooper
        if ($permittedAAACooperKeys.Count -gt 0) {
            $textboxResults.AppendText("Querying AAA Cooper...`r`n"); $mainForm.Refresh(); $rates=@{};
            foreach ($tariff in ($permittedAAACooperKeys.Keys|Sort)){
                try {
                    $keyData=$permittedAAACooperKeys[$tariff] 
                    if(!$keyData.APIToken -or !$keyData.CustomerNumber -or !$keyData.WhoAmI){throw "AAA Cooper creds missing from key file."}
                    
                    $aaaCooperShipmentData = [PSCustomObject]@{
                        OriginCity = $originCity; OriginState = $originState; OriginZip = $originZip; OriginCountryCode = $optionalShipmentDetails.OriginCountryCode 
                        DestinationCity = $destCity; DestinationState = $destState; DestinationZip = $destZip; DestinCountryCode = $optionalShipmentDetails.DestinationCountryCode 
                        BillDate = (Get-Date -Format 'MMddyyyy')
                        PrePaidCollect = "P" # Forced
                        RateEstimateRequestLine = $commoditiesForApi # This is the array of commodity objects
                        TotalPalletCountForPayload = ($commoditiesForApi | Where-Object {$_.HandlingUnitType -eq "Pallets"} | Measure-Object -Sum pieces).Sum # Example: sum pieces if pallets
                        PrimaryAccCode = if (($commoditiesForApi | Where-Object {$_.HandlingUnitType -eq "Pallets"}).Count -gt 0) { "PALET" } else { $null }
                    }

                    if(Get-Command Invoke-AAACooperApi -EA SilentlyContinue){
                        $cost=Invoke-AAACooperApi -KeyData $keyData -ShipmentData $aaaCooperShipmentData
                        if($cost){$rates[$tariff]=$cost}
                    } else { throw "Invoke-AAACooperApi missing."}
                } catch { $errMessage = if ($_.Exception) {$_.Exception.Message} else {$_}; $resultsText.AppendLine("  AAA Cooper ($tariff): Err - $errMessage") }
            };
            $lowest=Get-MinimumRate -RateResults $rates;
            if($lowest){
                $data=$permittedAAACooperKeys[$lowest.TariffName]; $margin=$Global:DefaultMarginPercentage
                if($data.MarginPercent){try{$margin=[double]$data.MarginPercent}catch{}}
                $quote=Calculate-QuotePrice -LowestCarrierCost $lowest.Cost -OriginZip $originZip -DestinationZip $destZip -Weight $totalWeight -FreightClass $firstClass -MarginPercent $margin
                if($quote.FinalPrice){
                    $resultsText.AppendLine(("  {0,-20}: {1,-15} (Tariff: {2})" -f "AAA Cooper", $quote.FinalPrice.ToString("C"), $lowest.TariffName))
                    $finalQuotes+=$quote
                    Write-QuoteToHistory -Carrier 'AAACooper' -Tariff $lowest.TariffName -OriginZip $originZip -DestinationZip $destZip -Weight $totalWeight -FreightClass $firstClass -LowestCost $lowest.Cost -FinalQuotedPrice $quote.FinalPrice -QuoteTimestamp $quoteTimestamp -OriginZipFull $originZip -DestinationZipFull $destZip
                } else { $resultsText.AppendLine("  AAA Cooper: Price Calc Err") }
            } else { $resultsText.AppendLine("  AAA Cooper: No Rates") }
        } else { $resultsText.AppendLine("  AAA Cooper: No Keys") }

    } finally { $VerbosePreference = $CurrentVerbosePreference }
    $resultsText.AppendLine("="*58)|Out-Null; $resultsText.AppendLine("* Prices are estimates.*")|Out-Null; $resultsText.AppendLine("--- End Quote ---")|Out-Null; $textboxResults.Text = $resultsText.ToString(); $statusBar.Text = "Quote complete."
})

# --- Settings Tab Handlers ---
$settingsCarrierChangedHandler = {
    param($sender, $e)
    if ($sender.Checked) {
        $custName = $null; if($comboBoxSelectCustomer_Settings.SelectedIndex -ge 0){$custName=$comboBoxSelectCustomer_Settings.SelectedItem.ToString()}
        if($custName -and $custName -ne "No Customers Found" -and $custName -ne "Select Customer"){ # Added check for placeholder items
            $carrier="Central" 
            if($radioSAIA.Checked){$carrier="SAIA"}
            elseif($radioRL.Checked){$carrier="RL"}
            elseif($radioAveritt.Checked){$carrier="Averitt"}
            elseif($radioAAACooper.Checked){$carrier="AAACooper"} 
            try{Populate-TariffListBox -SelectedCarrier $carrier -ListBoxControl $listBoxTariffs -LabelControl $labelSelectedTariff -ButtonControl $buttonSetMargin -TextboxControl $textBoxNewMargin -SelectedCustomerName $custName -AllCustomerProfiles $script:allCustomerProfiles}
            catch{$statusBar.Text="Err refresh settings: $($_.Exception.Message)"}
        } else {
            $listBoxTariffs.Items.Clear();$listBoxTariffs.Items.Add("Select Customer")
            $labelSelectedTariff.Text="Selected: (None)";$buttonSetMargin.Enabled=$false;$textBoxNewMargin.Enabled=$false;$textBoxNewMargin.Clear()
        }
    }
}
$radioCentral.Add_CheckedChanged($settingsCarrierChangedHandler)
$radioSAIA.Add_CheckedChanged($settingsCarrierChangedHandler)
$radioRL.Add_CheckedChanged($settingsCarrierChangedHandler)
$radioAveritt.Add_CheckedChanged($settingsCarrierChangedHandler)
$radioAAACooper.Add_CheckedChanged($settingsCarrierChangedHandler) 

$listBoxTariffs.Add_SelectedIndexChanged({ if($listBoxTariffs.SelectedIndex -ge 0){$itemTxt=$listBoxTariffs.SelectedItem.ToString();$tariffName=$null;$idx=$itemTxt.LastIndexOf('%');if($idx -gt 0){$sIdx=$itemTxt.LastIndexOf(' ',$idx);if($sIdx -gt 0){$tariffName=$itemTxt.Substring(0,$sIdx).Trim()}else{$tariffName=$itemTxt.Split(' ')[0].Trim()}}else{$tariffName=$itemTxt.Trim()};$labelSelectedTariff.Text="Selected: $tariffName";$textBoxNewMargin.Enabled=$true;$buttonSetMargin.Enabled=$true}else{$labelSelectedTariff.Text="Selected: (None)";$textBoxNewMargin.Enabled=$false;$buttonSetMargin.Enabled=$false} })

$buttonSetMargin.Add_Click({
    param($sender, $e)
    if($listBoxTariffs.SelectedIndex -lt 0){[System.Windows.Forms.MessageBox]::Show("Select tariff.");return}
    $itemTxt=$listBoxTariffs.SelectedItem.ToString();$tariffName=$null;$idx=$itemTxt.LastIndexOf('%');if($idx -gt 0){$sIdx=$itemTxt.LastIndexOf(' ',$idx);if($sIdx -gt 0){$tariffName=$itemTxt.Substring(0,$sIdx).Trim()}else{$tariffName=$itemTxt.Split(' ')[0].Trim()}}else{$tariffName=$itemTxt.Trim()}; if([string]::IsNullOrWhiteSpace($tariffName) -or $tariffName -match "^(No permitted|Select)"){ [System.Windows.Forms.MessageBox]::Show("Invalid tariff selected."); return };
    $marginInput=$textBoxNewMargin.Text;$marginPercent=$null;try{$mVal=[double]$marginInput;if($mVal -ge 0 -and $mVal -lt 100){$marginPercent=[math]::Round($mVal,1)}else{throw "Margin must be 0.0-99.9."}}catch{[System.Windows.Forms.MessageBox]::Show("Invalid margin %. Must be a number between 0.0 and 99.9.");return};
    $carrier="Central";if($radioSAIA.Checked){$carrier="SAIA"}elseif($radioRL.Checked){$carrier="RL"}elseif($radioAveritt.Checked){$carrier="Averitt"}elseif($radioAAACooper.Checked){$carrier="AAACooper"} 

    $keysPath=$null;$allKeys=@{};
    switch($carrier){
        "Central"{$keysPath=$script:centralKeysFolderPath;$allKeys=$script:allCentralKeys}
        "SAIA"{$keysPath=$script:saiaKeysFolderPath;$allKeys=$script:allSAIAKeys}
        "RL"{$keysPath=$script:rlKeysFolderPath;$allKeys=$script:allRLKeys}
        "Averitt"{$keysPath=$script:averittKeysFolderPath;$allKeys=$script:allAverittKeys}
        "AAACooper"{$keysPath=$script:aaaCooperKeysFolderPath;$allKeys=$script:allAAACooperKeys} 
        default{[System.Windows.Forms.MessageBox]::Show("Cannot find folder for carrier: $carrier");return}
    }
    $statusBar.Text="Updating...";$success=$false;
    try{if(!(Get-Command Update-TariffMargin -EA SilentlyContinue)){throw "Update func missing."};$success=Update-TariffMargin -TariffName $tariffName -AllKeysHashtable $allKeys -KeysFolderPath $keysPath -NewMarginPercent $marginPercent}catch{[System.Windows.Forms.MessageBox]::Show("Update Error: $($_.Exception.Message)");$statusBar.Text="Update failed.";return};
    if($success){$statusBar.Text="Updated '$tariffName'.";[System.Windows.Forms.MessageBox]::Show("Margin updated for $tariffName.");$custName=$null;if($comboBoxSelectCustomer_Settings.SelectedIndex -ge 0){$custName=$comboBoxSelectCustomer_Settings.SelectedItem.ToString()};if($custName -and $custName -ne "No Customers Found" -and $custName -ne "Select Customer"){Populate-TariffListBox -SelectedCarrier $carrier -ListBoxControl $listBoxTariffs -LabelControl $labelSelectedTariff -ButtonControl $buttonSetMargin -TextboxControl $textBoxNewMargin -SelectedCustomerName $custName -AllCustomerProfiles $script:allCustomerProfiles};$textBoxNewMargin.Clear()}else{$statusBar.Text="Update failed.";[System.Windows.Forms.MessageBox]::Show("Update failed for $tariffName.")}
})

# --- Reports Tab Handlers ---
$comboBoxReportType_SelectedIndexChanged_ScriptBlock = {
    param($sender, $e)
    $report = $null
    if ($comboBoxReportType.SelectedIndex -ge 0) {
        $report = $comboBoxReportType.SelectedItem.ToString()
    } else {
        $groupBoxReportCarrierSelect.Visible = $false
        $groupBoxReportTariffSelect.Visible = $false
        $labelReportTariff2.Visible = $false
        $listBoxReportTariff2.Visible = $false
        $labelSelectCsv.Visible = $false
        $textboxCsvPath.Visible = $false
        $buttonSelectCsv.Visible = $false
        $groupBoxReportAspInput.Visible = $false 
        $checkBoxApplyMargins.Visible = $false
        $buttonRunReport.Enabled = $false 
        return
    }

    $groupBoxReportCarrierSelect.Visible = $false
    $groupBoxReportTariffSelect.Visible = $false
    $labelReportTariff2.Visible = $false
    $listBoxReportTariff2.Visible = $false
    $labelSelectCsv.Visible = $false
    $textboxCsvPath.Visible = $false
    $buttonSelectCsv.Visible = $false
    $groupBoxReportAspInput.Visible = $false 
    $checkBoxApplyMargins.Visible = $false
    $buttonRunReport.Enabled = $true 
    
    $dynY = $reportsPanelY + (2 * $reportsControlSpacing) 

    $needsCarrierSelection = $report -in @("Carrier Comparison", "Avg Required Margin", "Required Margin for ASP")
    $needsTariffSelection = $needsCarrierSelection 
    
    $needsCsvInput = $report -in @("Carrier Comparison", "Avg Required Margin", "Required Margin for ASP", "Cross-Carrier ASP Analysis", "Margins by History")
    $needsManualASPInput = $report -in @("Required Margin for ASP", "Cross-Carrier ASP Analysis")
    $needsApplyMarginsCheckbox = $report -in @("Cross-Carrier ASP Analysis", "Margins by History")

    if ($needsCarrierSelection) {
        $groupBoxReportCarrierSelect.Location = [System.Drawing.Point]::new($reportsPanelX, $dynY)
        $groupBoxReportCarrierSelect.Visible = $true
        $dynY += $groupBoxReportCarrierSelect.Height + $verticalPaddingBetweenControls
    }

    if ($needsTariffSelection) {
        $groupBoxReportTariffSelect.Location = [System.Drawing.Point]::new($reportsPanelX, $dynY)
        $groupBoxReportTariffSelect.Visible = $true
        
        $needsTwoTariffs = $report -in @("Carrier Comparison", "Avg Required Margin")
        $labelReportTariff2.Visible = $needsTwoTariffs
        $listBoxReportTariff2.Visible = $needsTwoTariffs
        if ($needsTwoTariffs) {
            $labelReportTariff1.Text = "Tariff 1 (Base):"
        } else {
            $labelReportTariff1.Text = "Select Tariff:"
        }
        
        $custName = $null
        if ($comboBoxSelectCustomer_Reports.SelectedIndex -ge 0) {
            $custName = $comboBoxSelectCustomer_Reports.SelectedItem.ToString()
        }
        if ($custName -and $custName -ne "No Customers Found" -and $custName -ne "Select Customer") { 
            $carrierForTariffList = "Central" 
            if ($groupBoxReportCarrierSelect.Visible) { 
                if($radioReportSAIA.Checked){$carrierForTariffList="SAIA"}
                elseif($radioReportRL.Checked){$carrierForTariffList="RL"}
                elseif($radioReportAveritt.Checked){$carrierForTariffList="Averitt"}
                elseif($radioReportAAACooper.Checked){$carrierForTariffList="AAACooper"}
            }
            try {
                Populate-ReportTariffListBoxes -SelectedCarrier $carrierForTariffList `
                                               -ReportType $report `
                                               -SelectedCustomerName $custName `
                                               -AllCustomerProfiles $script:allCustomerProfiles `
                                               -ListBox1 $listBoxReportTariff1 `
                                               -Label1 $labelReportTariff1 `
                                               -ListBox2 $listBoxReportTariff2 `
                                               -Label2 $labelReportTariff2
            } catch { $statusBar.Text = "Error populating report tariff lists: $($_.Exception.Message)" }
        } else {
            $listBoxReportTariff1.Items.Clear(); $listBoxReportTariff1.Items.Add("Select Customer")
            $listBoxReportTariff2.Items.Clear(); $listBoxReportTariff2.Items.Add("Select Customer")
        }
        $dynY += $groupBoxReportTariffSelect.Height + $verticalPaddingBetweenControls
    }

    if ($needsCsvInput) {
        $labelSelectCsv.Location = [System.Drawing.Point]::new($reportsPanelX, $dynY)
        $labelSelectCsv.Visible = $true
        $textboxCsvPath.Location = [System.Drawing.Point]::new(($reportsPanelX + $reportsLabelWidth + 5), ($dynY - 3))
        $textboxCsvPath.Visible = $true
        $buttonSelectCsv.Location = [System.Drawing.Point]::new(($reportsPanelX + $reportsLabelWidth + $reportsInputWidth + 10), ($dynY - 5))
        $buttonSelectCsv.Visible = $true
        $dynY += $reportsControlSpacing
    }

    if ($needsManualASPInput) {
        $groupBoxReportAspInput.Location = [System.Drawing.Point]::new($reportsPanelX, $dynY)
        $groupBoxReportAspInput.Visible = $true
        $dynY += $groupBoxReportAspInput.Height + $verticalPaddingBetweenControls
    } else {
        $groupBoxReportAspInput.Visible = $false 
    }

    if ($needsApplyMarginsCheckbox) {
        $checkBoxApplyMargins.Location = [System.Drawing.Point]::new($reportsPanelX, $dynY)
        $checkBoxApplyMargins.Visible = $true
        $dynY += $checkBoxApplyMargins.Height + $verticalPaddingBetweenControls
    }
    
    $buttonRunReport.Location = [System.Drawing.Point]::new(($reportsPanelX + $reportsLabelWidth + 5), $dynY)
    $dynY += $buttonRunReport.Height + $verticalPaddingBetweenControls
    
    $textboxReportResults.Location = [System.Drawing.Point]::new($reportsPanelX, $dynY)
    $textboxReportResults.Height = $reportsPanel.ClientSize.Height - $dynY - $verticalPaddingBetweenControls
}
$comboBoxReportType.Add_SelectedIndexChanged($comboBoxReportType_SelectedIndexChanged_ScriptBlock)

$reportCarrierChangedHandler = { param($sender, $e); if ($sender.Checked) { $comboBoxReportType_SelectedIndexChanged_ScriptBlock.Invoke($comboBoxReportType, [System.EventArgs]::Empty) } }
$radioReportCentral.Add_CheckedChanged($reportCarrierChangedHandler)
$radioReportSAIA.Add_CheckedChanged($reportCarrierChangedHandler)
$radioReportRL.Add_CheckedChanged($reportCarrierChangedHandler)
$radioReportAveritt.Add_CheckedChanged($reportCarrierChangedHandler)
$radioReportAAACooper.Add_CheckedChanged($reportCarrierChangedHandler) 

$buttonSelectCsv.Add_Click({ param($sender, $e); $csvPath = Select-CsvFile -DialogTitle "Select Report Data CSV" -InitialDirectory $script:shipmentDataFolderPath; if ($csvPath) { $textboxCsvPath.Text = $csvPath } })

$buttonRunReport.Add_Click({
    param($sender, $e)
    $textboxReportResults.Clear(); $statusBar.Text = "Starting report..."; $mainForm.Refresh()
    $selectedCustomerName=$null;if($comboBoxSelectCustomer_Reports.SelectedIndex -ge 0){$selectedCustomerName=$comboBoxSelectCustomer_Reports.SelectedItem.ToString()};if(!$selectedCustomerName -or $selectedCustomerName -eq "No Customers Found" -or $selectedCustomerName -eq "Select Customer"){[System.Windows.Forms.MessageBox]::Show("Please select a valid customer.","Input Error");return}
    $customerProfile=$null;if($script:allCustomerProfiles.ContainsKey($selectedCustomerName)){$customerProfile=$script:allCustomerProfiles[$selectedCustomerName]};if(!$customerProfile){[System.Windows.Forms.MessageBox]::Show("Customer profile data missing for '$selectedCustomerName'.","Profile Error");return}
    $brokerProfile=$script:currentUserProfile;if(!$brokerProfile){[System.Windows.Forms.MessageBox]::Show("Broker profile missing. Please re-login.","Auth Error");return}
    if($comboBoxReportType.SelectedIndex -lt 0){[System.Windows.Forms.MessageBox]::Show("Please select a report type.","Input Error");return};$selectedReportType=$comboBoxReportType.SelectedItem.ToString()
    
    $csvPathForReport = $textboxCsvPath.Text
    $needsCsvForThisReport = $selectedReportType -in @("Carrier Comparison", "Avg Required Margin", "Required Margin for ASP", "Cross-Carrier ASP Analysis", "Margins by History")
    if($needsCsvForThisReport -and (!$csvPathForReport -or !(Test-Path $csvPathForReport -PathType Leaf))){
        [System.Windows.Forms.MessageBox]::Show("Please select a valid CSV data file for this report.","Input Error"); return
    }

    $reportsFolder=$script:currentUserReportsFolder;$reportFunction=$null;$reportParams=@{};
    if ($brokerProfile.Username){$reportParams.Username=$brokerProfile.Username}else{$reportParams.Username="UnknownUser"}
    if ($needsCsvForThisReport) {$reportParams.CsvFilePath = $csvPathForReport}
    $reportParams.UserReportsFolder = $reportsFolder # Common parameter for all reports that save output

    $selectedCarrierForReport=$null;$key1DataForReport=$null;$key2DataForReport=$null;$allKeysForReport=$null
    
    try {
        if ($groupBoxReportCarrierSelect.Visible) {
            if($radioReportCentral.Checked){$selectedCarrierForReport="Central";$allKeysForReport=$script:allCentralKeys}
            elseif($radioReportSAIA.Checked){$selectedCarrierForReport="SAIA";$allKeysForReport=$script:allSAIAKeys}
            elseif($radioReportRL.Checked){$selectedCarrierForReport="RL";$allKeysForReport=$script:allRLKeys}
            elseif($radioReportAveritt.Checked){$selectedCarrierForReport="Averitt";$allKeysForReport=$script:allAverittKeys}
            elseif($radioReportAAACooper.Checked){$selectedCarrierForReport="AAACooper";$allKeysForReport=$script:allAAACooperKeys} 
            else{throw "No carrier selected in the carrier selection groupbox."};

            if($listBoxReportTariff1.SelectedIndex -lt 0 -or $listBoxReportTariff1.SelectedItem.ToString() -match "^(No|Select)"){throw "Please select Tariff 1."};
            $tariff1Name=$listBoxReportTariff1.SelectedItem.ToString();if(!$allKeysForReport.ContainsKey($tariff1Name)){throw "Data for Tariff 1 ('$tariff1Name') is missing."};$key1DataForReport=$allKeysForReport[$tariff1Name];
            
            if($listBoxReportTariff2.Visible){ # If second tariff list is visible, it's needed
                if($listBoxReportTariff2.SelectedIndex -lt 0 -or $listBoxReportTariff2.SelectedItem.ToString() -match "^(No|Select)"){throw "Please select Tariff 2."};
                $tariff2Name=$listBoxReportTariff2.SelectedItem.ToString();if($tariff1Name -eq $tariff2Name){throw "Tariff 1 and Tariff 2 cannot be the same."};
                if(!$allKeysForReport.ContainsKey($tariff2Name)){throw "Data for Tariff 2 ('$tariff2Name') is missing."};$key2DataForReport=$allKeysForReport[$tariff2Name]
            }
        }
        
        switch ($selectedReportType) {
            "Carrier Comparison"{if(!$selectedCarrierForReport -or !$key1DataForReport -or !$key2DataForReport){throw "Carrier Comparison report requires a carrier and two distinct tariffs to be selected."};$reportFunction="Run-${selectedCarrierForReport}ComparisonReportGUI";$reportParams.Key1Data=$key1DataForReport;$reportParams.Key2Data=$key2DataForReport}
            "Avg Required Margin"{if(!$selectedCarrierForReport -or !$key1DataForReport -or !$key2DataForReport){throw "Avg Required Margin report requires a carrier and two distinct tariffs."};$reportFunction="Run-${selectedCarrierForReport}MarginReportGUI";$reportParams.BaseKeyData=$key1DataForReport;$reportParams.ComparisonKeyData=$key2DataForReport}
            "Required Margin for ASP"{if(!$selectedCarrierForReport -or !$key1DataForReport){throw "Required Margin for ASP report requires a carrier and one tariff."};if(!$groupBoxReportAspInput.Visible){throw "ASP Input groupbox is not visible for this report type, which is unexpected."};$aspInput=$textBoxDesiredAsp.Text;if(!($aspInput -match '^\d+(\.\d+)?$' -and ([decimal]$aspInput -gt 0))){throw "Invalid Desired Average Selling Price. Please enter a positive number."};$reportFunction="Calculate-${selectedCarrierForReport}MarginForASPReportGUI";$reportParams.CostAccountInfo=$key1DataForReport;$reportParams.DesiredASP=[decimal]$aspInput}
            "Cross-Carrier ASP Analysis"{if(!$groupBoxReportAspInput.Visible){throw "ASP Input groupbox is not visible for this report type, which is unexpected."};$aspInput=$textBoxDesiredAsp.Text;if(!($aspInput -match '^\d+(\.\d+)?$' -and ([decimal]$aspInput -gt 0))){throw "Invalid Desired Average Selling Price. Please enter a positive number."};$reportFunction="Run-CrossCarrierASPAnalysisGUI";$reportParams=@{BrokerProfile=$brokerProfile;SelectedCustomerProfile=$customerProfile;ReportsBaseFolder=$script:reportsBaseFolderPath;UserReportsFolder=$reportsFolder;AllCentralKeys=$script:allCentralKeys;AllSAIAKeys=$script:allSAIAKeys;AllRLKeys=$script:allRLKeys;AllAverittKeys=$script:allAverittKeys;AllAAACooperKeys=$script:allAAACooperKeys;DesiredASPValue=[decimal]$aspInput;ApplyMargins=$checkBoxApplyMargins.Checked;ASPFromHistory=$false;CsvFilePath=$csvPathForReport}}
            "Margins by History"{$reportFunction="Run-MarginsByHistoryAnalysisGUI";$reportParams=@{BrokerProfile=$brokerProfile;SelectedCustomerProfile=$customerProfile;ReportsBaseFolder=$script:reportsBaseFolderPath;UserReportsFolder=$reportsFolder;AllCentralKeys=$script:allCentralKeys;AllSAIAKeys=$script:allSAIAKeys;AllRLKeys=$script:allRLKeys;AllAverittKeys=$script:allAverittKeys;AllAAACooperKeys=$script:allAAACooperKeys;CsvFilePath=$csvPathForReport;ApplyMargins=$checkBoxApplyMargins.Checked}}
            default{throw "Selected report type '$selectedReportType' is not recognized or handled."}
        }
        
        $statusBar.Text="Running: $selectedReportType...";$mainForm.Refresh();Write-Verbose "Calling $reportFunction with parameters: $($reportParams | Out-String)";
        $reportOutputPath = & $reportFunction @reportParams
        
        if($reportOutputPath -and (Test-Path $reportOutputPath)){$textboxReportResults.Text="Report Complete!`nPath: $reportOutputPath";$statusBar.Text="Report Complete.";if([System.Windows.Forms.MessageBox]::Show("Report generated successfully.`n$reportOutputPath`n`nOpen report file?","Report Complete",[System.Windows.Forms.MessageBoxButtons]::YesNo,[System.Windows.Forms.MessageBoxIcon]::Information) -eq 'Yes'){try{Open-FileExplorer -Path $reportOutputPath}catch{[System.Windows.Forms.MessageBox]::Show("Failed to open report file: $($_.Exception.Message)","File Open Error")}}}
        else{$textboxReportResults.Text="Report failed to generate or no output path was returned.";$statusBar.Text="Report Failed.";[System.Windows.Forms.MessageBox]::Show("Report generation failed or did not produce an output file. Check console for warnings/errors.","Report Failed",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Warning)}
    } catch { 
        $errorMsg="Error running report '$selectedReportType':`n$($_.Exception.Message)";
        Write-Error $errorMsg
        $textboxReportResults.Text=$errorMsg
        $statusBar.Text="Report Error."
        [System.Windows.Forms.MessageBox]::Show($errorMsg,"Report Execution Error",[System.Windows.Forms.MessageBoxButtons]::OK,[System.Windows.Forms.MessageBoxIcon]::Error) 
    }
})

# --- Form Load/Close ---
$mainForm.Add_Shown({ if ($loginPanel.Visible) { $textboxUsername.Focus() } })
# Add the event handler to the ComboBox after the script block is defined
$comboBoxReportType.Add_SelectedIndexChanged($comboBoxReportType_SelectedIndexChanged_ScriptBlock)
[void]$mainForm.ShowDialog()
Write-Host "GUI Closed."
