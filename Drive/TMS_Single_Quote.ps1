# TMS_Single_Quote.ps1
# Module for handling single shipment quotes.

# Relies on functions from TMS_Helpers.ps1 and variables set by the main entry script (TMS_GUI.ps1)
# after loading TMS_Config.ps1.

function Run-SingleQuote {
    param(
        # --- Parameters from User/System Context ---
        [Parameter(Mandatory=$true)]
        [string]$Username,
        [Parameter(Mandatory=$true)]
        [hashtable]$UserConfig,         # Logged-in user's profile
        [Parameter(Mandatory=$true)]
        [hashtable]$AllCentralKeys,
        [Parameter(Mandatory=$true)]
        [hashtable]$AllSAIAKeys,
        [Parameter(Mandatory=$true)]
        [hashtable]$AllRLKeys,
        [Parameter(Mandatory=$true)]
        [hashtable]$AllAverittKeys,

        # --- Parameters passed from the GUI Input Form ---
        [Parameter(Mandatory=$true)]
        [string]$OriginZipParam,
        [Parameter(Mandatory=$true)]
        [string]$OriginCityParam,
        [Parameter(Mandatory=$true)]
        [string]$OriginStateParam,
        [Parameter(Mandatory=$true)]
        [string]$DestinationZipParam,
        [Parameter(Mandatory=$true)]
        [string]$DestinationCityParam,
        [Parameter(Mandatory=$true)]
        [string]$DestinationStateParam,
        [Parameter(Mandatory=$true)]
        [decimal]$WeightParam,
        [Parameter(Mandatory=$true)]
        [string]$FreightClassParam, # GUI validates this as a string representing a number

        # Optional parameters with defaults (GUI provides these, but good to have defaults)
        [int]$PiecesParam = 1,
        [double]$ItemLengthParam = 48.0,
        [double]$ItemWidthParam = 40.0,
        [double]$ItemHeightParam = 40.0,
        [string]$PackagingTypeParam = "PLT",
        [string]$DescriptionParam = "Freight",
        [decimal]$DeclaredValueParam = 0.0
    )

    # Ensure helper functions exist
    $requiredHelpers = @(
        "Clear-HostAndDrawHeader", "Get-PermittedKeys", "Invoke-CentralTransportApi",
        "Invoke-SAIAApi", "Invoke-RLApi", "Invoke-AverittApi", "Get-MinimumRate",
        "Calculate-QuotePrice", "Write-QuoteToHistory"
    )
    foreach ($helper in $requiredHelpers) {
        if (-not (Get-Command $helper -ErrorAction SilentlyContinue)) {
            Write-Error "FATAL ERROR in Run-SingleQuote: Required helper function '$helper' not found. Ensure TMS_Helpers.ps1 is correctly loaded."
            Read-Host "Press Enter to return..." # Keep console open if run directly
            return
        }
    }
    $useLoadingBar = Get-Command Write-LoadingBar -ErrorAction SilentlyContinue

    # Use parameters instead of Read-Host
    $originZip = $OriginZipParam
    $originCity = $OriginCityParam
    $originState = $OriginStateParam.ToUpper() # Ensure uppercase
    $destinationZip = $DestinationZipParam
    $destinationCity = $DestinationCityParam
    $destinationState = $DestinationStateParam.ToUpper() # Ensure uppercase
    $weight = $WeightParam # Already decimal
    $freightClass = $FreightClassParam # String, as passed from GUI

    # Optional details from parameters
    $pieces = $PiecesParam
    $itemLength = $ItemLengthParam
    $itemWidth = $ItemWidthParam
    $itemHeight = $ItemHeightParam
    $packagingType = $PackagingTypeParam
    $description = $DescriptionParam
    $declaredValue = $DeclaredValueParam
    
    # This object is used by R+L and SAIA's Invoke-Api functions for their optional parameters.
    # Averitt's payload is constructed more directly.
    $optionalShipmentDetailsForRLSAIA = [PSCustomObject]@{
        OriginCity = $originCity; OriginState = $originState; DestinationCity = $destinationCity; DestinationState = $destinationState;
        ItemWidth = $itemWidth; ItemHeight = $itemHeight; ItemLength = $itemLength; DeclaredValue = $declaredValue;
        Pieces = $pieces; PackagingType = $packagingType; Description = $description;
        CustomerData = $null; QuoteType = 'Domestic'; # Defaults for R+L if not otherwise specified
    }

    # Clear console if running in console mode (GUI calls this, but good for direct test)
    if (-not $Host.UI.RawUI.GetType().FullName.Contains("VisualStudio")) { # Avoid clearing VSCode terminal if possible
        Clear-HostAndDrawHeader -Title "Single Shipment Quote Results" -User $Username
    } else {
         Write-Host "`n--- Single Shipment Quote Results (User: $Username) ---"
    }

    Write-Host "`nFetching rates for: $originCity, $originState $originZip -> $destinationCity, $destinationState $destinationZip, Weight: $weight lbs, Class: $freightClass" -ForegroundColor Cyan

    # --- 2. Get Permitted Keys for the Logged-in User ---
    $permittedCentralKeys = Get-PermittedKeys -AllKeys $AllCentralKeys -AllowedKeyNames $UserConfig.AllowedCentralKeys
    $permittedSAIAKeys = Get-PermittedKeys -AllKeys $AllSAIAKeys -AllowedKeyNames $UserConfig.AllowedSAIAKeys
    $permittedRLKeys = Get-PermittedKeys -AllKeys $AllRLKeys -AllowedKeyNames $UserConfig.AllowedRLKeys
    $permittedAverittKeys = Get-PermittedKeys -AllKeys $AllAverittKeys -AllowedKeyNames $UserConfig.AllowedAverittKeys

    # --- 3. Fetch Rates for Each Carrier ---
    $centralRates = @{}; $saiaRates = @{}; $rlRates = @{}; $averittRates = @{}
    $quoteTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

    # Central Transport Rates
    if ($permittedCentralKeys.Count -gt 0) {
        Write-Host "`nQuerying Central Transport..." -ForegroundColor Gray; $idx = 0
        foreach ($tariffFileName in ($permittedCentralKeys.Keys | Sort-Object)) {
            $idx++; if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($idx * 100 / $permittedCentralKeys.Count)) -Message "Querying Central: $tariffFileName" }
            $keyData = $permittedCentralKeys[$tariffFileName]
            $cost = Invoke-CentralTransportApi -ApiKey $keyData.accessCode `
                                               -OriginZip $originZip -DestinationZip $destinationZip `
                                               -Weight $weight -FreightClass $freightClass `
                                               -customerNumber $keyData.customerNumber
            if ($cost -ne $null) { $centralRates[$tariffFileName] = $cost }
        }
        if ($useLoadingBar) { Write-Progress -Activity "Querying Central Tariffs" -Completed }
    } else { Write-Warning "No permitted Central Transport keys for user $Username." }

    # SAIA Rates
    if ($permittedSAIAKeys.Count -gt 0) {
         Write-Host "`nQuerying SAIA..." -ForegroundColor Gray; $idx = 0
        foreach ($tariffFileName in ($permittedSAIAKeys.Keys | Sort-Object)) {
             $idx++; if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($idx * 100 / $permittedSAIAKeys.Count)) -Message "Querying SAIA: $tariffFileName" }
             $keyData = $permittedSAIAKeys[$tariffFileName]
             $cost = Invoke-SAIAApi -OriginZip $originZip -DestinationZip $destinationZip `
                                    -OriginCity $originCity -OriginState $originState `
                                    -DestinationCity $destinationCity -DestinationState $destinationState `
                                    -Weight $weight -Class $freightClass `
                                    -KeyData $keyData -Details $optionalShipmentDetailsForRLSAIA 
             if ($cost -ne $null) { $saiaRates[$tariffFileName] = $cost }
        }
        if ($useLoadingBar) { Write-Progress -Activity "Querying SAIA Tariffs" -Completed }
    } else { Write-Warning "No permitted SAIA keys for user $Username." }

    # R+L Rates
    if ($permittedRLKeys.Count -gt 0) {
         Write-Host "`nQuerying R+L Carriers..." -ForegroundColor Gray; $idx = 0
        foreach ($tariffFileName in ($permittedRLKeys.Keys | Sort-Object)) {
             $idx++; if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($idx * 100 / $permittedRLKeys.Count)) -Message "Querying R+L: $tariffFileName" }
             $keyData = $permittedRLKeys[$tariffFileName]
             $cost = Invoke-RLApi -OriginZip $originZip -DestinationZip $destinationZip `
                                  -Weight $weight -Class $freightClass `
                                  -KeyData $keyData -ShipmentDetails $optionalShipmentDetailsForRLSAIA
             if ($cost -ne $null) { $rlRates[$tariffFileName] = $cost }
        }
        if ($useLoadingBar) { Write-Progress -Activity "Querying R+L Tariffs" -Completed }
    } else { Write-Warning "No permitted R+L keys for user $Username." }

    # Averitt Rates
    if ($permittedAverittKeys.Count -gt 0) {
        Write-Host "`nQuerying Averitt..." -ForegroundColor Gray; $idx = 0
        foreach ($tariffFileName in ($permittedAverittKeys.Keys | Sort-Object)) {
            $idx++; if ($useLoadingBar) { Write-LoadingBar -PercentComplete ([int]($idx * 100 / $permittedAverittKeys.Count)) -Message "Querying Averitt: $tariffFileName" }
            $keyData = $permittedAverittKeys[$tariffFileName]

            $averittShipmentPayload = [PSCustomObject]@{
                ServiceLevel = "STND"; PaymentTerms = "PPD"; PaymentPayer = "Shipper"
                PickupDate   = (Get-Date).ToString("yyyyMMdd")
                OriginAccount = if ($keyData.ContainsKey('OriginAccount')) { $keyData.OriginAccount } else { "000000" } 
                OriginCity = $originCity; OriginStateProvince = $originState; OriginPostalCode = $originZip; OriginCountry = "USA"
                DestinationAccount = if ($keyData.ContainsKey('DestinationAccount')) { $keyData.DestinationAccount } else { "000000" } 
                DestinationCity = $destinationCity; DestinationStateProvince = $destinationState; DestinationPostalCode = $destinationZip; DestinationCountry = "USA"
                BillToAccount = if ($keyData.ContainsKey('BillToAccount')) { $keyData.BillToAccount } else { "000000" } 
                BillToName = if ($keyData.ContainsKey('BillToName')) { $keyData.BillToName } else { "TMS User" }
                BillToAddress = if ($keyData.ContainsKey('BillToAddress')) { $keyData.BillToAddress } else { "123 Main St" } 
                BillToCity = if ($keyData.ContainsKey('BillToCity')) { $keyData.BillToCity } else { $originCity } 
                BillToStateProvince = if ($keyData.ContainsKey('BillToStateProvince')) { $keyData.BillToStateProvince } else { $originState }
                BillToPostalCode = if ($keyData.ContainsKey('BillToPostalCode')) { $keyData.BillToPostalCode } else { $originZip }
                BillToCountry = if ($keyData.ContainsKey('BillToCountry')) { $keyData.BillToCountry } else { "USA" }
                Commodity1_Classification = $freightClass; Commodity1_Weight = $weight; Commodity1_Pieces = $pieces
                Commodity1_Length = $itemLength; Commodity1_Width = $itemWidth; Commodity1_Height = $itemHeight
                Commodity1_PackagingType = $packagingType; Commodity1_Description = $description; Commodity1_Stackable = "Y" 
                Commodity2_Classification = $null; Commodity2_Weight = $null; Commodity2_Pieces = $null 
                Commodity3_Classification = $null; Commodity3_Weight = $null; Commodity3_Pieces = $null 
                Commodity4_Classification = $null; Commodity4_Weight = $null; Commodity4_Pieces = $null 
                Commodity5_Classification = $null; Commodity5_Weight = $null; Commodity5_Pieces = $null 
                AccessorialCodes = ""; HazardousContactName = ""; HazardousContactPhone = ""
            }
            
            $cost = Invoke-AverittApi -KeyData $keyData -ShipmentData $averittShipmentPayload
            if ($cost -ne $null) { $averittRates[$tariffFileName] = $cost }
        }
        if ($useLoadingBar) { Write-Progress -Activity "Querying Averitt Tariffs" -Completed }
    } else { Write-Warning "No permitted Averitt keys for user $Username." }

    # --- 4. Calculate Best Price PER CARRIER and Display Quote ---
    Write-Host "`n"
    Write-Host "==================== SHIPMENT QUOTE (Console Output) ====================" -ForegroundColor Cyan
    Write-Host "Quote Date: $($quoteTimestamp)"
    Write-Host "Generated By: $($Username)"
    Write-Host "--------------------------------------------------------" -ForegroundColor Gray
    Write-Host "Origin:      $originCity, $originState $originZip"
    Write-Host "Destination: $destinationCity, $destinationState $destinationZip"
    Write-Host "Weight:      $($weight) lbs, Class: $($freightClass), Pieces: $pieces"
    if ($itemLength -ne 48.0 -or $itemWidth -ne 40.0 -or $itemHeight -ne 40.0 -or $packagingType -ne "PLT" -or $description -ne "Freight") {
        Write-Host "Details:     $($itemLength)Lx$($itemWidth)Wx$($itemHeight)H per piece, Pkg: $packagingType, Desc: $description"
    }
    if ($declaredValue -gt 0) { Write-Host "Declared Val:$($declaredValue.ToString("C"))" }
    Write-Host "--------------------------------------------------------" -ForegroundColor Gray
    Write-Host "Carrier Options (Price to Customer):" -ForegroundColor Yellow

    $finalQuotes = @() 

    function Process-CarrierQuote {
        param (
            [string]$CarrierDisplayName,
            [hashtable]$PermittedCarrierKeys,
            [hashtable]$CarrierRates, 
            [ref]$FinalQuotesRef      
        )
        $lowestRateInfo = Get-MinimumRate -RateResults $CarrierRates
        if ($lowestRateInfo -ne $null) {
            $lowestTariffData = $PermittedCarrierKeys[$lowestRateInfo.TariffName] 
            $marginToUse = $Global:DefaultMarginPercentage 
            if ($lowestTariffData -ne $null -and $lowestTariffData.ContainsKey('MarginPercent')) {
                 try { $marginToUse = [double]$lowestTariffData.MarginPercent } catch {
                     Write-Warning "Invalid MarginPercent for $CarrierDisplayName tariff '$($lowestRateInfo.TariffName)'. Using default."
                 }
            }
            $quoteDetails = Calculate-QuotePrice -LowestCarrierCost $lowestRateInfo.Cost `
                                                 -OriginZip $originZip -DestinationZip $destinationZip `
                                                 -Weight ([double]$weight) -FreightClass $freightClass ` # Ensure weight is double for Calculate-QuotePrice
                                                 -MarginPercent $marginToUse
            if ($quoteDetails.FinalPrice -ne $null) {
                 Write-Host ("  {0,-20}: {1,-15} (Tariff: {2})" -f $CarrierDisplayName, $quoteDetails.FinalPrice.ToString("C"), $lowestRateInfo.TariffName) -ForegroundColor Green
                 $FinalQuotesRef.Value += [PSCustomObject]@{Carrier=$CarrierDisplayName; Tariff=$lowestRateInfo.TariffName; Price=$quoteDetails.FinalPrice; Cost=$lowestRateInfo.Cost}
                 Write-QuoteToHistory -Carrier $CarrierDisplayName -TariffName $lowestRateInfo.TariffName `
                                      -OriginZip $originZip -DestinationZip $destinationZip -Weight ([double]$weight) -FreightClass $freightClass `
                                      -LowestCost $lowestRateInfo.Cost -FinalQuotedPrice $quoteDetails.FinalPrice -QuoteTimestamp $quoteTimestamp
            } else { Write-Host "  $($CarrierDisplayName.PadRight(20)): Error calculating final price for tariff '$($lowestRateInfo.TariffName)'." -ForegroundColor Red }
        } else { Write-Host "  $($CarrierDisplayName.PadRight(20)): No valid rates found." -ForegroundColor Gray }
    }

    Process-CarrierQuote -CarrierDisplayName "Central Transport" -PermittedCarrierKeys $permittedCentralKeys -CarrierRates $centralRates -FinalQuotesRef ([ref]$finalQuotes)
    Process-CarrierQuote -CarrierDisplayName "SAIA" -PermittedCarrierKeys $permittedSAIAKeys -CarrierRates $saiaRates -FinalQuotesRef ([ref]$finalQuotes)
    Process-CarrierQuote -CarrierDisplayName "R+L Carriers" -PermittedCarrierKeys $permittedRLKeys -CarrierRates $rlRates -FinalQuotesRef ([ref]$finalQuotes)
    Process-CarrierQuote -CarrierDisplayName "Averitt" -PermittedCarrierKeys $permittedAverittKeys -CarrierRates $averittRates -FinalQuotesRef ([ref]$finalQuotes)

    Write-Host "========================================================" -ForegroundColor Cyan
    Write-Host "* Prices are estimates and subject to verification. *" -ForegroundColor Gray
    Write-Host "`n--- End of Quote (Console Output) ---" -ForegroundColor Yellow
}

Write-Host "Single Quote module loaded." -ForegroundColor Cyan
