# TMS_GUI.ps1
# Main script to launch the TMS GUI Application. This is the primary entry point.

# --- Determine Script Root and Scope it ---
$script:scriptRoot = $null # Initialize with script scope
if ($PSScriptRoot) { $script:scriptRoot = $PSScriptRoot }
elseif ($MyInvocation.MyCommand.Path) { $script:scriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent }
else {
    Write-Error "FATAL: Cannot determine script root directory. GUI cannot start."
    Read-Host "Press Enter to exit..."
    exit 1
}
Write-Verbose "Script Root Determined for GUI: '$script:scriptRoot'"

# --- Dot-Source Module Files ---
Write-Verbose "Starting GUI: Loading core TMS modules..."
try {
    $configPath = Join-Path $script:scriptRoot "TMS_Config.ps1"; . $configPath
    Write-Verbose "GUI: Successfully dot-sourced TMS_Config.ps1."
    $moduleFiles = @(
        "TMS_Helpers.ps1", "TMS_Auth.ps1", "TMS_Carrier_Central.ps1", "TMS_Carrier_SAIA.ps1",
        "TMS_Carrier_RL.ps1", "TMS_Carrier_Averitt.ps1", "TMS_Reports.ps1",
        "TMS_Single_Quote.ps1", "TMS_Settings.ps1"
    )
    foreach ($moduleFile in $moduleFiles) {
        $modulePath = Join-Path $script:scriptRoot $moduleFile
        if (-not (Test-Path $modulePath -PathType Leaf)) { throw "GUI FATAL: Module '$moduleFile' not found." }
        . $modulePath; Write-Verbose "GUI: Successfully dot-sourced '$moduleFile'."
    }
     Write-Verbose "GUI: All core TMS modules loaded successfully."
} catch {
    Write-Error "GUI: FATAL ERROR loading modules: $($_.Exception.Message)"; Read-Host "Press Enter..."; exit 1
}

# --- Global Variables for GUI State ---
$Global:currentUserProfile = $null; $Global:allCentralKeys = $null; $Global:allSAIAKeys = $null
$Global:allRLKeys = $null; $Global:allAverittKeys = $null; $Global:currentUserReportsFolder = $null
$Global:customerProfiles = $null

# --- Resolve Full Paths for Data Folders ---
$script:centralKeysFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $script:defaultCentralKeysFolderName
$script:saiaKeysFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $script:defaultSAIAKeysFolderName
$script:rlKeysFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $script:defaultRLKeysFolderName
$script:averittKeysFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $script:defaultAverittKeysFolderName
$script:userAccountsFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $script:defaultUserAccountsFolderName
$script:customerAccountsFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $script:defaultCustomerAccountsFolderName
$script:reportsBaseFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $script:defaultReportsBaseFolderName
$script:shipmentDataFolderPath = Join-Path -Path $script:scriptRoot -ChildPath $script:defaultShipmentDataFolderName

# --- Ensure Required Base Data Folders Exist ---
Write-Verbose "GUI: Ensuring base data directories exist..."
Ensure-DirectoryExists -Path $script:centralKeysFolderPath; Ensure-DirectoryExists -Path $script:saiaKeysFolderPath
Ensure-DirectoryExists -Path $script:rlKeysFolderPath; Ensure-DirectoryExists -Path $script:averittKeysFolderPath
Ensure-DirectoryExists -Path $script:userAccountsFolderPath; Ensure-DirectoryExists -Path $script:customerAccountsFolderPath
Ensure-DirectoryExists -Path $script:reportsBaseFolderPath; Ensure-DirectoryExists -Path $script:shipmentDataFolderPath
Write-Verbose "GUI: Base directory check complete."

# --- Pre-load All Carrier Keys and Customer Profiles ---
Write-Host "`nInitializing TMS GUI Application..." -ForegroundColor Yellow
$Global:allCentralKeys = Load-KeysFromFolder -KeysFolderPath $script:centralKeysFolderPath -CarrierName "Central Transport"
$Global:allSAIAKeys = Load-KeysFromFolder -KeysFolderPath $script:saiaKeysFolderPath -CarrierName "SAIA"
$Global:allRLKeys = Load-KeysFromFolder -KeysFolderPath $script:rlKeysFolderPath -CarrierName "RL Carriers"
$Global:allAverittKeys = Load-KeysFromFolder -KeysFolderPath $script:averittKeysFolderPath -CarrierName "Averitt"
Write-Host "Carrier Key loading complete."
$Global:customerProfiles = Load-AllCustomerProfiles -CustomerAccountsFolderPath $script:customerAccountsFolderPath
Write-Host "Customer profile loading complete. $($Global:customerProfiles.Count) profiles loaded."

Add-Type -AssemblyName System.Windows.Forms; Add-Type -AssemblyName System.Drawing

# --- Login Window Function (Restored Full Version with Fix) ---
function Show-LoginWindow {
    Write-Host "DEBUG (GUI): Show-LoginWindow (Full) called."
    $loginForm = New-Object System.Windows.Forms.Form
    $loginForm.Text = "TMS Login"
    $loginForm.Size = New-Object System.Drawing.Size(320, 230) 
    $loginForm.StartPosition = "CenterScreen"
    $loginForm.FormBorderStyle = 'FixedDialog'
    $loginForm.MaximizeBox = $false
    $loginForm.MinimizeBox = $false
    $loginForm.CancelButton = $null 

    $labelUser = New-Object System.Windows.Forms.Label; $labelUser.Text = "Username:"; $labelUser.Location = New-Object System.Drawing.Point(20, 20); $labelUser.Size = New-Object System.Drawing.Size(80, 20); $loginForm.Controls.Add($labelUser)
    $textUser = New-Object System.Windows.Forms.TextBox; $textUser.Location = New-Object System.Drawing.Point(100, 20); $textUser.Size = New-Object System.Drawing.Size(180, 20); $textUser.TabIndex = 0; $loginForm.Controls.Add($textUser)
    $labelPass = New-Object System.Windows.Forms.Label; $labelPass.Text = "Password:"; $labelPass.Location = New-Object System.Drawing.Point(20, 55); $labelPass.Size = New-Object System.Drawing.Size(80, 20); $loginForm.Controls.Add($labelPass)
    $textPass = New-Object System.Windows.Forms.TextBox; $textPass.Location = New-Object System.Drawing.Point(100, 55); $textPass.Size = New-Object System.Drawing.Size(180, 20); $textPass.PasswordChar = '*'; $textPass.TabIndex = 1; $loginForm.Controls.Add($textPass)
    $messageLabel = New-Object System.Windows.Forms.Label; $messageLabel.Location = New-Object System.Drawing.Point(20, 90); $messageLabel.Size = New-Object System.Drawing.Size(260, 40); $messageLabel.ForeColor = [System.Drawing.Color]::Red; $loginForm.Controls.Add($messageLabel)
    $loginButton = New-Object System.Windows.Forms.Button; $loginButton.Text = "Login"; $loginButton.Location = New-Object System.Drawing.Point(110, 145); $loginButton.Size = New-Object System.Drawing.Size(100, 30); $loginButton.TabIndex = 2; $loginForm.AcceptButton = $loginButton; $loginForm.Controls.Add($loginButton)
    
    $loginButton.Add_Click({
        Write-Host "DEBUG (GUI): Login button clicked. Username: '$($textUser.Text)'" 
        $username = $textUser.Text; $passwordPlainText = $textPass.Text
        if ([string]::IsNullOrWhiteSpace($username) -or [string]::IsNullOrWhiteSpace($passwordPlainText)) {
            $messageLabel.Text = "Username and Password are required."; Write-Host "DEBUG (GUI): Username or Password empty."; return 
        }
        Write-Host "DEBUG (GUI): Calling Authenticate-User. UserAccountsFolderPath: '$($script:userAccountsFolderPath)'" 
        $Global:currentUserProfile = Authenticate-User -Username $username -PasswordPlainText $passwordPlainText -UserAccountsFolderPath $script:userAccountsFolderPath
        if ($Global:currentUserProfile -ne $null) {
            Write-Host "DEBUG (GUI): Authentication successful for '$($Global:currentUserProfile.Username)'" -ForegroundColor Green 
            $Global:currentUserReportsFolder = Join-Path $script:reportsBaseFolderPath $Global:currentUserProfile.Username
            Ensure-DirectoryExists -Path $Global:currentUserReportsFolder
            $loginForm.DialogResult = [System.Windows.Forms.DialogResult]::OK 
            $loginForm.Close() 
        } else {
            Write-Host "DEBUG (GUI): Authentication failed for '$username'." -ForegroundColor Red 
            $messageLabel.Text = "Login failed. Check username/password."; $textPass.Text = ""; $textPass.Focus()
        }
    })

    Write-Host "DEBUG (GUI): Showing full login form..." 
    $dialogResult = $loginForm.ShowDialog()
    Write-Host "DEBUG (GUI): Full login form closed. DialogResult from ShowDialog(): '$($dialogResult)'."
    Write-Host "DEBUG (GUI): After full ShowDialog(), currentUserProfile.Username: $($Global:currentUserProfile.Username)"

    if ($Global:currentUserProfile -ne $null -and $dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        Write-Host "DEBUG (GUI): Show-LoginWindow (Full) returning TRUE." ; return $true 
    } else {
        Write-Host "DEBUG (GUI): Show-LoginWindow (Full) returning FALSE. currentUserProfile is null or DialogResult was not OK ('$($dialogResult)')." 
        $Global:currentUserProfile = $null; return $false 
    }
}


# --- Function to Display Single Quote Input View in a Panel ---
function Display-SingleQuoteViewInPanel {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Panel]$ParentPanel
    )
    Write-Host "DEBUG (GUI): Display-SingleQuoteViewInPanel called."
    $ParentPanel.Controls.Clear() 

    # $controls hashtable is useful for programmatic access if needed, but for click handler, direct find is more robust
    # $script:tempControls = @{} # Using script scope for helper, or pass $controls as [ref]

    $yPos = 20
    $labelWidth = 150
    $textWidth = 220
    $leftMargin = 20
    $textLeftMargin = $leftMargin + $labelWidth + 10
    
    function Add-PanelLabeledTextBox {
        param($panel, $controlName, $labelText, $initialValue, [ref]$currentYPos, $toolTipText=$null)
        $label = New-Object System.Windows.Forms.Label; $label.Text = $labelText; $label.Location = New-Object System.Drawing.Point($leftMargin, $currentYPos.Value); $label.Size = New-Object System.Drawing.Size($labelWidth, 20); $panel.Controls.Add($label)
        $textBox = New-Object System.Windows.Forms.TextBox; $textBox.Name = $controlName; $textBox.Text = $initialValue; $textBox.Location = New-Object System.Drawing.Point($textLeftMargin, $currentYPos.Value); $textBox.Size = New-Object System.Drawing.Size($textWidth, 20)
        if ($toolTipText) { $toolTip = New-Object System.Windows.Forms.ToolTip; $toolTip.SetToolTip($textBox, $toolTipText) }
        $panel.Controls.Add($textBox); $currentYPos.Value += 30
        # $script:tempControls[$controlName] = $textBox # Store if needed for other dynamic access
        return $textBox # Return for direct assignment if preferred, though not strictly needed if finding by name
    }

    Add-PanelLabeledTextBox $ParentPanel "textOriginZip" "Origin ZIP (5 digits):" "" ([ref]$yPos) -toolTipText "Enter 5-digit origin ZIP code."
    Add-PanelLabeledTextBox $ParentPanel "textOriginCity" "Origin City:" "" ([ref]$yPos)
    Add-PanelLabeledTextBox $ParentPanel "textOriginState" "Origin State (2 letters):" "" ([ref]$yPos) -toolTipText "e.g., TX, CA"
    $yPos += 5 
    Add-PanelLabeledTextBox $ParentPanel "textDestZip" "Destination ZIP (5 digits):" "" ([ref]$yPos)
    Add-PanelLabeledTextBox $ParentPanel "textDestCity" "Destination City:" "" ([ref]$yPos)
    Add-PanelLabeledTextBox $ParentPanel "textDestState" "Destination State (2 letters):" "" ([ref]$yPos)
    $yPos += 5
    Add-PanelLabeledTextBox $ParentPanel "textWeight" "Total Weight (lbs):" "" ([ref]$yPos)
    Add-PanelLabeledTextBox $ParentPanel "textFreightClass" "Freight Class (50-500):" "70" ([ref]$yPos)
    $yPos += 5
    Add-PanelLabeledTextBox $ParentPanel "textPieces" "Number of Pieces (opt):" "1" ([ref]$yPos)
    Add-PanelLabeledTextBox $ParentPanel "textLength" "Length/piece (in, opt):" "48" ([ref]$yPos)
    Add-PanelLabeledTextBox $ParentPanel "textWidth" "Width/piece (in, opt):" "40" ([ref]$yPos)
    Add-PanelLabeledTextBox $ParentPanel "textHeight" "Height/piece (in, opt):" "40" ([ref]$yPos)
    Add-PanelLabeledTextBox $ParentPanel "textPkgType" "Packaging (PLT, CTN, opt):" "PLT" ([ref]$yPos)
    Add-PanelLabeledTextBox $ParentPanel "textDesc" "Description (opt):" "Freight" ([ref]$yPos)
    Add-PanelLabeledTextBox $ParentPanel "textDeclaredValue" "Declared Value (USD, opt):" "0.00" ([ref]$yPos)
    
    $runQuoteButton = New-Object System.Windows.Forms.Button; $runQuoteButton.Text = "Get Quote"; 
    $runQuoteButton.Size = New-Object System.Drawing.Size(140, 35); 
    $panelWidthInt = [int]$ParentPanel.Width
    $buttonWidthInt = [int]$runQuoteButton.Width
    $runQuoteButton.Location = New-Object System.Drawing.Point( (($panelWidthInt - $buttonWidthInt) / 2), ($yPos + 15) ) 
    $runQuoteButton.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    
    $runQuoteButton.Add_Click({
        # Retrieve controls directly from the ParentPanel by name inside the click event
        $originZipCtrl = $ParentPanel.Controls.Find("textOriginZip", $true)[0]
        $originCityCtrl = $ParentPanel.Controls.Find("textOriginCity", $true)[0]
        $originStateCtrl = $ParentPanel.Controls.Find("textOriginState", $true)[0]
        $destZipCtrl = $ParentPanel.Controls.Find("textDestZip", $true)[0]
        $destCityCtrl = $ParentPanel.Controls.Find("textDestCity", $true)[0]
        $destStateCtrl = $ParentPanel.Controls.Find("textDestState", $true)[0]
        $weightCtrl = $ParentPanel.Controls.Find("textWeight", $true)[0]
        $freightClassCtrl = $ParentPanel.Controls.Find("textFreightClass", $true)[0]
        $piecesCtrl = $ParentPanel.Controls.Find("textPieces", $true)[0]
        $lengthCtrl = $ParentPanel.Controls.Find("textLength", $true)[0]
        $widthCtrl = $ParentPanel.Controls.Find("textWidth", $true)[0]
        $heightCtrl = $ParentPanel.Controls.Find("textHeight", $true)[0]
        $pkgTypeCtrl = $ParentPanel.Controls.Find("textPkgType", $true)[0]
        $descCtrl = $ParentPanel.Controls.Find("textDesc", $true)[0]
        $declaredValueCtrl = $ParentPanel.Controls.Find("textDeclaredValue", $true)[0]

        # Validation using retrieved controls
        if ($null -eq $originZipCtrl) { [System.Windows.Forms.MessageBox]::Show("Internal error: OriginZip control could not be found on panel.", "GUI Error", "OK", "Error"); return }
        if (-not ($originZipCtrl.Text -match '^\d{5}$')) { [System.Windows.Forms.MessageBox]::Show("Invalid Origin ZIP.", "Validation Error", "OK", "Error"); $originZipCtrl.Focus(); return }
        if ([string]::IsNullOrWhiteSpace($originCityCtrl.Text)) { [System.Windows.Forms.MessageBox]::Show("Origin City is required.", "Validation Error", "OK", "Error"); $originCityCtrl.Focus(); return }
        if (-not ($originStateCtrl.Text -match '^[A-Za-z]{2}$')) { [System.Windows.Forms.MessageBox]::Show("Invalid Origin State (e.g., TX).", "Validation Error", "OK", "Error"); $originStateCtrl.Focus(); return }
        if (-not ($destZipCtrl.Text -match '^\d{5}$')) { [System.Windows.Forms.MessageBox]::Show("Invalid Destination ZIP.", "Validation Error", "OK", "Error"); $destZipCtrl.Focus(); return }
        if ([string]::IsNullOrWhiteSpace($destCityCtrl.Text)) { [System.Windows.Forms.MessageBox]::Show("Destination City is required.", "Validation Error", "OK", "Error"); $destCityCtrl.Focus(); return }
        if (-not ($destStateCtrl.Text -match '^[A-Za-z]{2}$')) { [System.Windows.Forms.MessageBox]::Show("Invalid Destination State (e.g., CA).", "Validation Error", "OK", "Error"); $destStateCtrl.Focus(); return }
        
        $weightVal = 0.0; if (-not [decimal]::TryParse($weightCtrl.Text, [ref]$weightVal) -or $weightVal -le 0) { [System.Windows.Forms.MessageBox]::Show("Invalid Weight. Must be a positive number.", "Validation Error", "OK", "Error"); $weightCtrl.Focus(); return }
        $classValText = $freightClassCtrl.Text; $classValNum = 0.0; 
        if (-not [double]::TryParse($classValText, [ref]$classValNum) -or $classValNum -lt 50 -or $classValNum -gt 500) { 
            [System.Windows.Forms.MessageBox]::Show("Invalid Freight Class. Must be a number between 50 and 500 (e.g., 77.5).", "Validation Error", "OK", "Error"); $freightClassCtrl.Focus(); return 
        }

        $piecesVal = 1; if (-not [int]::TryParse($piecesCtrl.Text, [ref]$piecesVal) -or $piecesVal -le 0) { $piecesVal = 1 }
        $lengthVal = 48.0; if (-not [double]::TryParse($lengthCtrl.Text, [ref]$lengthVal)) { $lengthVal = 48.0 }
        $widthVal = 40.0; if (-not [double]::TryParse($widthCtrl.Text, [ref]$widthVal)) { $widthVal = 40.0 }
        $heightVal = 40.0; if (-not [double]::TryParse($heightCtrl.Text, [ref]$heightVal)) { $heightVal = 40.0 }
        $pkgTypeVal = "PLT"; if (-not [string]::IsNullOrWhiteSpace($pkgTypeCtrl.Text)) { $pkgTypeVal = $pkgTypeCtrl.Text.Trim().ToUpper() }
        $descVal = "Freight"; if (-not [string]::IsNullOrWhiteSpace($descCtrl.Text)) { $descVal = $descCtrl.Text.Trim() }
        $decValueVal = 0.0; if (-not [decimal]::TryParse($declaredValueCtrl.Text, [ref]$decValueVal)) { $decValueVal = 0.0 }

        [System.Windows.Forms.MessageBox]::Show("Processing quote... Please check the PowerShell console for results.", "Processing Quote", "OK", "Information")
        $ParentPanel.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
        $runQuoteButton.Enabled = $false

        Run-SingleQuote -Username $Global:currentUserProfile.Username `
                        -UserConfig $Global:currentUserProfile `
                        -AllCentralKeys $Global:allCentralKeys `
                        -AllSAIAKeys $Global:allSAIAKeys `
                        -AllRLKeys $Global:allRLKeys `
                        -AllAverittKeys $Global:allAverittKeys `
                        -OriginZipParam $originZipCtrl.Text `
                        -OriginCityParam $originCityCtrl.Text `
                        -OriginStateParam $originStateCtrl.Text.ToUpper() `
                        -DestinationZipParam $destZipCtrl.Text `
                        -DestinationCityParam $destCityCtrl.Text `
                        -DestinationStateParam $destStateCtrl.Text.ToUpper() `
                        -WeightParam $weightVal `
                        -FreightClassParam $classValText ` 
                        -PiecesParam $piecesVal `
                        -ItemLengthParam $lengthVal `
                        -ItemWidthParam $widthVal `
                        -ItemHeightParam $heightVal `
                        -PackagingTypeParam $pkgTypeVal `
                        -DescriptionParam $descVal `
                        -DeclaredValueParam $decValueVal
        
        Write-Host "`nSingle quote process from GUI input finished. Results are in the console." -ForegroundColor Yellow
        $ParentPanel.Cursor = [System.Windows.Forms.Cursors]::Default
        $runQuoteButton.Enabled = $true
    })
    $ParentPanel.Controls.Add($runQuoteButton)
}

# --- Function to Display Default Dashboard View ---
function Display-DashboardView {
    param(
        [Parameter(Mandatory=$true)]
        [System.Windows.Forms.Panel]$ParentPanel
    )
    $ParentPanel.Controls.Clear()
    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = "TMS Dashboard Area - Select an option from the menu."
    $infoLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Regular)
    $infoLabel.TextAlign = "MiddleCenter"
    $infoLabel.Dock = "Fill"
    $ParentPanel.Controls.Add($infoLabel)
}


# --- Main Application Window Function ---
function Show-MainWindow {
    $mainForm = New-Object System.Windows.Forms.Form
    $mainForm.Text = "Transportation Management System - Welcome $($Global:currentUserProfile.Username)"
    $mainForm.Size = New-Object System.Drawing.Size(850, 650) 
    $mainForm.StartPosition = "CenterScreen"
    $mainForm.WindowState = "Normal" 

    $Global:mainContentPanel = New-Object System.Windows.Forms.Panel 
    $mainContentPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $mainContentPanel.BackColor = [System.Drawing.Color]::WhiteSmoke 
    
    $menuStrip = New-Object System.Windows.Forms.MenuStrip
    $menuStrip.Dock = [System.Windows.Forms.DockStyle]::Top 

    # File Menu
    $fileMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("&File") 
    $dashboardMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("&Dashboard")
    $dashboardMenuItem.Add_Click({ Display-DashboardView -ParentPanel $mainContentPanel })
    $logoutMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("&Logout")
    $logoutMenuItem.Add_Click({ 
        $Global:currentUserProfile = $null; $mainForm.Close(); Write-Host "User logged out. Restarting login..."
        Start-Sleep -Seconds 1; if (Show-LoginWindow) { Show-MainWindow } else { [System.Windows.Forms.Application]::Exit() }
    })
    $exitMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("E&xit")
    $exitMenuItem.Add_Click({ [System.Windows.Forms.Application]::Exit() })
    $fileMenuItem.DropDownItems.AddRange(@($dashboardMenuItem, $logoutMenuItem, (New-Object System.Windows.Forms.ToolStripSeparator), $exitMenuItem))
    $menuStrip.Items.Add($fileMenuItem)

    # Quote Menu
    $quoteMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("&Quoting")
    $singleQuoteMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("&Single Shipment Quote")
    $singleQuoteMenuItem.Add_Click({
        Display-SingleQuoteViewInPanel -ParentPanel $mainContentPanel
    })
    $quoteMenuItem.DropDownItems.Add($singleQuoteMenuItem)
    $menuStrip.Items.Add($quoteMenuItem)

    # Reports Menu
    $reportsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("&Reports")
    $carrierReportsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Carrier &Specific Reports")
        $centralReportsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("&Central Transport"); $centralReportsMenuItem.Add_Click({ Display-DashboardView -ParentPanel $mainContentPanel; Show-CarrierReportWindow -CarrierName "Central" -AllCarrierKeys $Global:allCentralKeys -UserConfig $Global:currentUserProfile -UserReportsFolder $Global:currentUserReportsFolder})
        $saiaReportsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("&SAIA"); $saiaReportsMenuItem.Add_Click({ Display-DashboardView -ParentPanel $mainContentPanel; Show-CarrierReportWindow -CarrierName "SAIA" -AllCarrierKeys $Global:allSAIAKeys -UserConfig $Global:currentUserProfile -UserReportsFolder $Global:currentUserReportsFolder})
        $rlReportsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("R+&L Carriers"); $rlReportsMenuItem.Add_Click({ Display-DashboardView -ParentPanel $mainContentPanel; Show-CarrierReportWindow -CarrierName "RL" -AllCarrierKeys $Global:allRLKeys -UserConfig $Global:currentUserProfile -UserReportsFolder $Global:currentUserReportsFolder})
        $averittReportsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("&Averitt"); $averittReportsMenuItem.Add_Click({ Display-DashboardView -ParentPanel $mainContentPanel; Show-CarrierReportWindow -CarrierName "Averitt" -AllCarrierKeys $Global:allAverittKeys -UserConfig $Global:currentUserProfile -UserReportsFolder $Global:currentUserReportsFolder})
    $carrierReportsMenuItem.DropDownItems.AddRange(@($centralReportsMenuItem, $saiaReportsMenuItem, $rlReportsMenuItem, $averittReportsMenuItem))
    $reportsMenuItem.DropDownItems.Add($carrierReportsMenuItem)
    $crossCarrierASPMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("&Cross-Carrier ASP Analysis"); $crossCarrierASPMenuItem.Add_Click({ Display-DashboardView -ParentPanel $mainContentPanel; if(Get-Command Run-CrossCarrierASPAnalysis -EA SilentlyContinue){ [System.Windows.Forms.MessageBox]::Show("Cross-Carrier ASP Analysis will run in console.", "Console Interaction", "OK", "Information"); Run-CrossCarrierASPAnalysis -UserProfile $Global:currentUserProfile -ReportsBaseFolder $script:reportsBaseFolderPath -AllCentralKeys $Global:allCentralKeys -AllSAIAKeys $Global:allSAIAKeys -AllRLKeys $Global:allRLKeys -AllAverittKeys $Global:allAverittKeys; Write-Host "`nCross-Carrier ASP finished." -FG Yellow } else { [System.Windows.Forms.MessageBox]::Show("Function not found.", "Error")} })
    $reportsMenuItem.DropDownItems.Add($crossCarrierASPMenuItem)
    $manageMyReportsItem = New-Object System.Windows.Forms.ToolStripMenuItem("&Manage My Reports"); $manageMyReportsItem.Add_Click({ Display-DashboardView -ParentPanel $mainContentPanel; if (Get-Command Manage-UserReports -EA SilentlyContinue) { Manage-UserReports -UserReportsFolder $Global:currentUserReportsFolder } else {[System.Windows.Forms.MessageBox]::Show("Function not found.", "Error")} })
    $reportsMenuItem.DropDownItems.Add($manageMyReportsItem)
    $menuStrip.Items.Add($reportsMenuItem)
    
    # Settings Menu
    $settingsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("S&ettings")
    $manageMarginsMenuItem = New-Object System.Windows.Forms.ToolStripMenuItem("Manage Tariff &Margins"); $manageMarginsMenuItem.Add_Click({ Display-DashboardView -ParentPanel $mainContentPanel; if(Get-Command Show-SettingsMenu -EA SilentlyContinue){  [System.Windows.Forms.MessageBox]::Show("Settings will run in console.", "Console Interaction", "OK", "Information"); Show-SettingsMenu -UserProfile $Global:currentUserProfile -AllCentralKeys $Global:allCentralKeys -AllSAIAKeys $Global:allSAIAKeys -AllRLKeys $Global:allRLKeys -AllAverittKeys $Global:allAverittKeys; Write-Host "`nSettings management finished." -FG Yellow } else { [System.Windows.Forms.MessageBox]::Show("Function not found.", "Error")} })
    $settingsMenuItem.DropDownItems.Add($manageMarginsMenuItem)
    $menuStrip.Items.Add($settingsMenuItem)

    $mainForm.Controls.Add($mainContentPanel) 
    $mainForm.Controls.Add($menuStrip) 
    
    $statusBar = New-Object System.Windows.Forms.StatusBar; $statusBar.Text = "Ready"; $mainForm.Controls.Add($statusBar)
    
    Display-DashboardView -ParentPanel $mainContentPanel
    
    $mainForm.ShowDialog() 
}

# --- Carrier Report Window Function (Generic for launching reports) ---
function Show-CarrierReportWindow {
    param(
        [string]$CarrierName,
        [hashtable]$AllCarrierKeys, 
        [hashtable]$UserConfig,     
        [string]$UserReportsFolder 
    )

    $reportForm = New-Object System.Windows.Forms.Form
    $reportForm.Text = "$CarrierName Reports"
    $reportForm.Size = New-Object System.Drawing.Size(520, 420) 
    $reportForm.StartPosition = "CenterParent"
    $reportForm.FormBorderStyle = 'FixedDialog'

    $yPos = 20
    $labelTitle = New-Object System.Windows.Forms.Label; $labelTitle.Text = "$CarrierName Report Options for $($UserConfig.Username):"; $labelTitle.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold); $labelTitle.Location = New-Object System.Drawing.Point(20, $yPos); $labelTitle.AutoSize=$true; $reportForm.Controls.Add($labelTitle); $yPos += 35

    $permittedKeys = $null
    $allowedKeysPropertyName = "Allowed${CarrierName}Keys" 
    if ($CarrierName -eq "RL") { $allowedKeysPropertyName = "AllowedRLKeys" } 

    if ($UserConfig.ContainsKey($allowedKeysPropertyName) -and $UserConfig.$allowedKeysPropertyName -is [array]) {
        $allowedKeyNames = $UserConfig.$allowedKeysPropertyName
        $permittedKeys = Get-PermittedKeys -AllKeys $AllCarrierKeys -AllowedKeyNames $allowedKeyNames
    }

    if ($null -eq $permittedKeys -or $permittedKeys.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("You do not have any permitted keys/tariffs for $CarrierName. Please check your user profile settings in the respective customer account file.", "Permissions Error", "OK", "Error")
        $reportForm.Dispose(); return
    }

    $groupBoxReportType = New-Object System.Windows.Forms.GroupBox; $groupBoxReportType.Text = "Select Report Type"; $groupBoxReportType.Location = New-Object System.Drawing.Point(20, $yPos); $groupBoxReportType.Size = New-Object System.Drawing.Size(460, 100); $reportForm.Controls.Add($groupBoxReportType); $yPos += 110
    
    $radioComparison = New-Object System.Windows.Forms.RadioButton; $radioComparison.Text = "Comparison Report"; $radioComparison.Location = New-Object System.Drawing.Point(15, 25); $radioComparison.AutoSize=$true; $radioComparison.Checked = $true; $groupBoxReportType.Controls.Add($radioComparison)
    $radioAvgMargin = New-Object System.Windows.Forms.RadioButton; $radioAvgMargin.Text = "Average Required Margin (Base vs. Comparison)"; $radioAvgMargin.Location = New-Object System.Drawing.Point(15, 50); $radioAvgMargin.AutoSize=$true; $groupBoxReportType.Controls.Add($radioAvgMargin)
    $radioMarginASP = New-Object System.Windows.Forms.RadioButton; $radioMarginASP.Text = "Required Margin for Desired ASP (Single Tariff)"; $radioMarginASP.Location = New-Object System.Drawing.Point(15, 75); $radioMarginASP.AutoSize=$true; $groupBoxReportType.Controls.Add($radioMarginASP)

    $labelKey1 = New-Object System.Windows.Forms.Label; $labelKey1.Text = "Tariff/Key 1 (or Base Key):"; $labelKey1.Location = New-Object System.Drawing.Point(20, $yPos); $labelKey1.AutoSize=$true; $reportForm.Controls.Add($labelKey1)
    $comboKey1 = New-Object System.Windows.Forms.ComboBox; $comboKey1.Location = New-Object System.Drawing.Point(250, $yPos - 3); $comboKey1.Size = New-Object System.Drawing.Size(230, 21); $comboKey1.DropDownStyle = "DropDownList"
    $permittedKeys.GetEnumerator() | Sort-Object Value.Name | ForEach-Object { $comboKey1.Items.Add($_.Value.Name) } 
    if($comboKey1.Items.Count -gt 0) { $comboKey1.SelectedIndex = 0 }; $reportForm.Controls.Add($comboKey1); $yPos += 30

    $labelKey2 = New-Object System.Windows.Forms.Label; $labelKey2.Text = "Tariff/Key 2 (for Comparison/Avg Margin):"; $labelKey2.Location = New-Object System.Drawing.Point(20, $yPos); $labelKey2.AutoSize=$true; $reportForm.Controls.Add($labelKey2)
    $comboKey2 = New-Object System.Windows.Forms.ComboBox; $comboKey2.Location = New-Object System.Drawing.Point(250, $yPos - 3); $comboKey2.Size = New-Object System.Drawing.Size(230, 21); $comboKey2.DropDownStyle = "DropDownList"
    $permittedKeys.GetEnumerator() | Sort-Object Value.Name | ForEach-Object { $comboKey2.Items.Add($_.Value.Name) }
    if($comboKey2.Items.Count -gt 1) { $comboKey2.SelectedIndex = 1 } elseif($comboKey2.Items.Count -gt 0) { $comboKey2.SelectedIndex = 0 }; $reportForm.Controls.Add($comboKey2); $yPos += 30
    
    $labelASP = New-Object System.Windows.Forms.Label; $labelASP.Text = "Desired ASP (for Margin for ASP report):"; $labelASP.Location = New-Object System.Drawing.Point(20, $yPos); $labelASP.AutoSize=$true; $reportForm.Controls.Add($labelASP)
    $textASP = New-Object System.Windows.Forms.TextBox; $textASP.Location = New-Object System.Drawing.Point(250, $yPos -3); $textASP.Size = New-Object System.Drawing.Size(100,20); $textASP.Text="0.00"; $reportForm.Controls.Add($textASP); $yPos += 30

    $buttonSelectCsv = New-Object System.Windows.Forms.Button; $buttonSelectCsv.Text = "Select CSV Data File..."; $buttonSelectCsv.Location = New-Object System.Drawing.Point(20, $yPos); $buttonSelectCsv.Size = New-Object System.Drawing.Size(210, 25); $reportForm.Controls.Add($buttonSelectCsv)
    $labelCsvPath = New-Object System.Windows.Forms.Label; $labelCsvPath.Location = New-Object System.Drawing.Point(250, $yPos + 4); $labelCsvPath.Size = New-Object System.Drawing.Size(230, 20); $labelCsvPath.Text = "No CSV selected"; $reportForm.Controls.Add($labelCsvPath); $yPos += 40
    $buttonSelectCsv.Add_Click({
        $dialogTitle = "Select Shipment Data CSV for $CarrierName Report"
        if ($CarrierName -eq "Averitt") { 
            $dialogTitle = "Select DETAILED Shipment CSV for Averitt (e.g., shipments.csv format)"
        }
        $selectedCsv = Select-CsvFile -DialogTitle $dialogTitle -InitialDirectory $script:shipmentDataFolderPath
        if ($selectedCsv) { $labelCsvPath.Text = (Split-Path $selectedCsv -Leaf); $labelCsvPath.Tag = $selectedCsv } 
    })
    
    $buttonRunReport = New-Object System.Windows.Forms.Button; $buttonRunReport.Text = "Run Report"; $buttonRunReport.Location = New-Object System.Drawing.Point(200, $yPos); $buttonRunReport.Size = New-Object System.Drawing.Size(120, 30); $reportForm.Controls.Add($buttonRunReport)
    $buttonRunReport.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $buttonRunReport.Add_Click({
        $selectedKey1Name = $comboKey1.SelectedItem
        $selectedKey2Name = $comboKey2.SelectedItem
        $csvPath = $labelCsvPath.Tag 

        if ([string]::IsNullOrWhiteSpace($selectedKey1Name)) { [System.Windows.Forms.MessageBox]::Show("Please select Tariff/Key 1.", "Input Error", "OK", "Error"); return }
        if ([string]::IsNullOrWhiteSpace($csvPath)) { [System.Windows.Forms.MessageBox]::Show("Please select a CSV data file.", "Input Error", "OK", "Error"); return }

        $key1Details = $permittedKeys.GetEnumerator() | Where-Object {$_.Value.Name -eq $selectedKey1Name} | Select-Object -First 1 -ExpandProperty Value
        $key2Details = $null
        if (-not [string]::IsNullOrWhiteSpace($selectedKey2Name)) {
            $key2Details = $permittedKeys.GetEnumerator() | Where-Object {$_.Value.Name -eq $selectedKey2Name} | Select-Object -First 1 -ExpandProperty Value
        }

        $reportPath = $null; $reportFunctionName = $null; $params = @{}

        if ($radioComparison.Checked) {
            if ($null -eq $key2Details) { [System.Windows.Forms.MessageBox]::Show("Please select Tariff/Key 2 for Comparison report.", "Input Error", "OK", "Error"); return }
            if ($key1Details.Name -eq $key2Details.Name) { [System.Windows.Forms.MessageBox]::Show("Tariff/Key 1 and Tariff/Key 2 must be different for Comparison.", "Input Error", "OK", "Error"); return }
            $reportFunctionName = "Run-${CarrierName}ComparisonReportGUI"
            $params = @{ Key1Data = $key1Details; Key2Data = $key2Details; CsvFilePath = $csvPath; Username = $UserConfig.Username; UserReportsFolder = $UserReportsFolder }
        } elseif ($radioAvgMargin.Checked) {
            if ($null -eq $key2Details) { [System.Windows.Forms.MessageBox]::Show("Please select Tariff/Key 2 for Avg Req Margin report (to compare against).", "Input Error", "OK", "Error"); return }
            if ($key1Details.Name -eq $key2Details.Name) { [System.Windows.Forms.MessageBox]::Show("Base Key and Comparison Key must be different.", "Input Error", "OK", "Error"); return }
            $reportFunctionName = "Run-${CarrierName}MarginReportGUI" 
            $params = @{ BaseKeyData = $key1Details; ComparisonKeyData = $key2Details; CsvFilePath = $csvPath; Username = $UserConfig.Username; UserReportsFolder = $UserReportsFolder }
        } elseif ($radioMarginASP.Checked) {
            $desiredASPValue = 0.0
            if (-not [decimal]::TryParse($textASP.Text, [ref]$desiredASPValue) -or $desiredASPValue -le 0) {
                [System.Windows.Forms.MessageBox]::Show("Please enter a valid positive Desired ASP value.", "Input Error", "OK", "Error"); return
            }
            $reportFunctionName = "Calculate-${CarrierName}MarginForASPReportGUI"
            $params = @{ CostAccountInfo = $key1Details; DesiredASP = $desiredASPValue; CsvFilePath = $csvPath; Username = $UserConfig.Username; UserReportsFolder = $UserReportsFolder }
        }

        if ($reportFunctionName -and (Get-Command $reportFunctionName -ErrorAction SilentlyContinue)) {
            Write-Host "GUI: Calling $reportFunctionName in background job..."
            $buttonRunReport.Enabled = $false; $reportForm.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
            
            $job = Start-Job -ScriptBlock {
                param($CmdName, $CmdParamsToJob, $JobScriptRoot)
                . (Join-Path $JobScriptRoot "TMS_Config.ps1") 
                . (Join-Path $JobScriptRoot "TMS_Helpers.ps1")
                . (Join-Path $JobScriptRoot "TMS_Carrier_Central.ps1") 
                . (Join-Path $JobScriptRoot "TMS_Carrier_SAIA.ps1")
                . (Join-Path $JobScriptRoot "TMS_Carrier_RL.ps1")
                . (Join-Path $JobScriptRoot "TMS_Carrier_Averitt.ps1")
                
                if (Get-Command $CmdName -ErrorAction SilentlyContinue) {
                    & $CmdName @CmdParamsToJob
                } else {
                    Write-Error "Job Error: Command '$CmdName' not found in job scope."
                    return $null 
                }
            } -ArgumentList $reportFunctionName, $params, $script:scriptRoot 
            
            Register-ObjectEvent -InputObject $job -EventName StateChanged -Action {
                param($SenderJob)
                if ($SenderJob.State -in ('Completed', 'Failed', 'Stopped')) {
                    Write-Host "GUI: Background report job '$($SenderJob.Name)' finished with state: $($SenderJob.State)"
                    $receivedReportPath = $null
                    $jobErrors = $null
                    try {
                        if ($SenderJob.Error.Count -gt 0) {
                            $jobErrors = ($SenderJob.Error | ForEach-Object {$_.ToString()}) -join [System.Environment]::NewLine
                        }
                        $receivedReportPath = Receive-Job -Job $SenderJob -Keep
                    } catch {
                        $jobErrors = "Error receiving job results: $($_.Exception.Message)"
                    }
                    
                    $reportForm.Invoke([Action[object, string, string, System.Windows.Forms.Button, System.Windows.Forms.Form]]{
                        param($jobState, $repPath, $errMsgs, $btnRun, $frmReport)

                        $frmReport.Cursor = [System.Windows.Forms.Cursors]::Default
                        $btnRun.Enabled = $true

                        if ($jobState -eq 'Failed' -or ($null -eq $repPath -and $jobState -ne 'Completed' -and $errMsgs)) { 
                             [System.Windows.Forms.MessageBox]::Show("Report generation failed or was cancelled. Check console for details. Errors: $($errMsgs)", "Report Error", "OK", "Error")
                        } elseif ($repPath) {
                            [System.Windows.Forms.MessageBox]::Show("Report generated successfully: $repPath", "Report Complete", "OK", "Information")
                        } else { 
                             [System.Windows.Forms.MessageBox]::Show("Report process completed, but no report file was generated or an issue occurred. Check console. Errors: $($errMsgs)", "Report Notice", "OK", "Warning")
                        }
                    }, @($SenderJob.State, $receivedReportPath, $jobErrors, $buttonRunReport, $reportForm))
                    
                    Remove-Job -Job $SenderJob -Force
                    Unregister-Event -SourceIdentifier $SenderJob.InstanceId.ToString() 
                }
            } -SourceIdentifier $job.InstanceId.ToString() 

            [System.Windows.Forms.MessageBox]::Show("Report generation started in the background. You will be notified upon completion.", "Report Running", "OK", "Information")

        } else {
            [System.Windows.Forms.MessageBox]::Show("Report function '$reportFunctionName' not found for carrier $CarrierName.", "Error", "OK", "Error")
        }
    })

    $reportForm.ShowDialog()
}


# --- Main Application Entry Point ---
Write-Host "TMS GUI is starting..." -ForegroundColor Green

if (Show-LoginWindow) { 
    Show-MainWindow
} else {
    Write-Warning "Login failed or cancelled. Exiting GUI."
}

Write-Host "TMS GUI application finished." -ForegroundColor Green
[System.Windows.Forms.Application]::Exit() 
