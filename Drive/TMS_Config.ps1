# TMS_Config.ps1
# Description: Configuration settings for the TMS Tool.

# --- API Endpoints ---
$script:centralApiUri = 'https://client-api.centraltransport.com/api/v1/Quote/byDimensions'
$script:saiaApiUri = "https://api.saia.com/rate-quote/webservice/ratequote/customer-api"
$script:rlApiUri = "http://api.rlcarriers.com/1.0.3/RateQuoteService.asmx"
$script:averittApiUri = "https://api.averittexpress.com/rate-quotes/dynamicpricing" 
$script:aaaCooperApiUri = "https://api2.aaacooper.com:8200/sapi30/wsGenEst"

# --- Default Folder Names (relative to script root) ---
$script:defaultUserAccountsFolderName = "user_accounts"
$script:defaultCustomerAccountsFolderName = "customer_accounts"
$script:defaultReportsBaseFolderName = "reports"
$script:defaultShipmentDataFolderName = "shipmentData"
$script:defaultHistoricalQuotesFileName = "historical_quotes_generated.csv" # Inside shipmentDataFolderName

# Carrier Specific Key/Tariff Folders
$script:defaultCentralKeysFolderName = "keys_central"
$script:defaultSAIAKeysFolderName = "keys_saia"
$script:defaultRLKeysFolderName = "keys_rl"
$script:defaultAverittKeysFolderName = "keys_averitt"
$script:defaultAAACooperKeysFolderName = "keys_aaacooper" # New Folder for AAA Cooper

# --- Default Settings ---
$script:DefaultMarginPercentage = 20.0 # Default margin if not specified in tariff file
$script:DefaultMinProfit = 50.0       # Default minimum profit if not specified

# --- Logging & Debug ---
$script:EnableFullApiLogging = $false # Set to $true to log full API requests/responses (can be verbose)
$script:ApiLogPath = Join-Path $PSScriptRoot "api_logs" # Ensure this directory exists if logging enabled

# --- GUI Appearance (Optional, can be overridden in TMS_GUI.ps1) ---
$script:guiFontFamily = "Segoe UI"
$script:guiFontSize = 9
$script:guiThemeColor = "0,120,215" # RGB for a blueish theme

Write-Verbose "TMS Configuration loaded."
