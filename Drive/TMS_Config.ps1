# TMS_Config.ps1
# Configuration for the TMS Tool (Margins stored in Key Files)

# --- Default Folder Names (Relative to script root) ---
# These can be overridden by parameters passed to the main entry script if needed.
$script:defaultCentralKeysFolderName = "keys_central"
$script:defaultSAIAKeysFolderName = "keys_saia"
$script:defaultRLKeysFolderName = "keys_rl"
$script:defaultAverittKeysFolderName = "keys_averitt" # <<< AVERITT ADDED >>>
$script:defaultUserAccountsFolderName = "user_accounts"     # For BROKER logins
$script:defaultCustomerAccountsFolderName = "customer_accounts" # For CUSTOMER profiles
$script:defaultReportsBaseFolderName = "reports"
$script:defaultShipmentDataFolderName = "shipmentData"

# --- API Endpoints ---
# Use $script: scope so helper functions and carrier modules can access them.
$script:centralApiUri = 'https://client-api.centraltransport.com/api/v1/Quote/byClass' # Example URL
$script:saiaApiUri = "https://api.saia.com/rate-quote/webservice/ratequote/customer-api" # Example URL
$script:rlApiUri = "http://api.rlcarriers.com/1.0.3/RateQuoteService.asmx" # Example URL
$script:averittApiUri = "https://api.averittexpress.com/rate-quotes/dynamicpricing" # <<< AVERITT ADDED >>>

# --- Pricing & Margin Configuration ---
# Default margin percentage to use ONLY if a key file is missing the 'MarginPercent' line.
# Use Global scope for values potentially needed across different script scopes if not passed directly.
$Global:DefaultMarginPercentage = 15.0 # e.g., 15%

# --- Historical Pricing Configuration ---
# Specify the filename of the CSV containing actual historical shipment data.
# Place this file inside the 'shipmentData' folder.
$Global:HistoricalDataSourceFileName = "shipmentHistory.csv"

# Filename for the log where THIS tool writes its generated quotes (separate from source data).
$Global:HistoricalQuotesLogFileName = "historical_quotes_generated.csv"

# Weight tolerance for historical matching (+/- this percentage).
$Global:HistoricalWeightTolerance = 0.10 # +/- 10%

# --- Other Settings ---
# Add any other global configuration variables here.

Write-Host "Configuration defaults loaded." -ForegroundColor Cyan
# Note: Actual folder paths are resolved in the main entry script (TMS_GUI.ps1) after this file is loaded.
