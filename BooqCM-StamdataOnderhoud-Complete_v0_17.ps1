# BooqCM-StamdataOnderhoud.ps1
# PowerShell ISE 5.1 compatible script voor onderhoud stamdata via Booq CM API
# Versie: 1.3 - Critical Bugfix Release: Encoding, Availability, SalesPoints, Forms

<#
.SYNOPSIS
    Interactief script voor het onderhouden van stamdata via de Booq C&M API
    
.DESCRIPTION
    Dit script biedt een interactieve interface voor het beheren van stamdata zoals:
    - Promoties
    - Time Periods
    - Availabilities
    - Customers
    
    Alle acties worden opgeslagen in sessies die later kunnen worden hervat.
    
.PARAMETER clientId
    OAuth2 Client ID
    
.PARAMETER clientSecret
    OAuth2 Client Secret
    
.PARAMETER environment
    Omgeving: 'sandbox' of 'production'
    
.PARAMETER impersonateClientId
    Optioneel: Client ID om te impersoneren
    
.PARAMETER saveOrgJson
    Sla originele JSON responses op
    
.PARAMETER onboarding
    Start in onboarding modus
    
.PARAMETER showExtraDetails
    Schakel extra details en API request logging in
    
.PARAMETER DebugMode
    Uitgebreide debug output
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$clientId,
    
    [Parameter(Mandatory=$true)]
    [string]$clientSecret,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet('sandbox', 'production')]
    [string]$environment = 'sandbox',
    
    [Parameter(Mandatory=$false)]
    [string]$impersonateClientId = $null,
    
    [Parameter(Mandatory=$false)]
    [bool]$saveOrgJson = $false,
    
    [Parameter(Mandatory=$false)]
    [switch]$onboarding,
    
    [Parameter(Mandatory=$false)]
    [switch]$showExtraDetails,
    
    [Parameter(Mandatory=$false)]
    [switch]$DebugMode
)

# ============================================================================
# INITIALISATIE
# ============================================================================

# Console encoding instellen voor correcte weergave van speciale karakters
# FIX v1.3: Check host type om errors in ISE/remote sessions te voorkomen
try {
    if ($Host.Name -eq "ConsoleHost") {
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
    }
    $PSDefaultParameterValues['*:Encoding'] = 'utf8'
}
catch {
    # Negeer encoding errors in ISE of andere hosts
    Write-Verbose "Console encoding kon niet worden ingesteld: $_"
}

# Start locatie bewaren
$script:startLocation = Get-Location
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import common module
$commonModulePath = Join-Path $scriptPath "BooqAPI-Common.psm1"
if (-not (Test-Path $commonModulePath)) {
    Write-Host "FOUT: BooqAPI-Common.psm1 niet gevonden in $scriptPath" -ForegroundColor Red
    exit 1
}

Import-Module $commonModulePath -Force

# Initialiseer logging in start locatie
$logPath = Initialize-Logging -ScriptName "BooqCM-Stamdata" -NetworkLocation $script:startLocation
Write-Log "=== Booq CM Stamdata Onderhoud gestart ===" -level INFO
Write-Log "Script versie: 1.3 (Critical Bugfix: Encoding, Availability, SalesPoints, Forms)" -level INFO
Write-Log "Omgeving: $environment" -level INFO
if ($showExtraDetails) {
    Write-Log "Extra details en API request logging: INGESCHAKELD" -level INFO
}
if ($saveOrgJson) {
    Write-Log "JSON response export: INGESCHAKELD" -level INFO
    # Maak JSON output directory aan
    $script:jsonOutputFolder = Join-Path $script:startLocation "CM-JSON-Responses"
    if (-not (Test-Path $script:jsonOutputFolder)) {
        New-Item -ItemType Directory -Path $script:jsonOutputFolder -Force | Out-Null
        Write-Log "JSON output folder aangemaakt: $script:jsonOutputFolder" -level INFO
    }
}

# ============================================================================
# API REQUEST LOGGING FUNCTIE
# ============================================================================

function Write-RequestLog {
    <#
    .SYNOPSIS
        Logt alle API requests naar een CSV bestand
    .DESCRIPTION
        Schrijft gedetailleerde informatie over elke API request naar CM-API-Requests.csv
        Alleen actief wanneer showExtraDetails parameter is ingeschakeld
    #>
    param(
        [hashtable]$RequestLogEntry
    )
    
    if ($showExtraDetails) {
        try {
            $logFilePath = Join-Path $script:startLocation "CM-API-Requests.csv"
            
            # Create file with headers if it doesn't exist
            if (-not (Test-Path $logFilePath)) {
                $csvHeaders = @(
                    "Timestamp",
                    "Method",
                    "URL",
                    "Status",
                    "StatusText",
                    "RequestHeaders",
                    "ResponseHeaders",
                    "RequestBody",
                    "ResponseBody",
                    "ContentType",
                    "Time",
                    "ServerIP"
                ) -join ";"
                
                $csvHeaders | Out-File -FilePath $logFilePath -Encoding UTF8
            }
            
            # Prepare CSV line with proper escaping
            $csvLine = @(
                $RequestLogEntry.Timestamp,
                $RequestLogEntry.Method,
                $RequestLogEntry.URL,
                $RequestLogEntry.Status,
                $RequestLogEntry.StatusText,
                ('"' + ($RequestLogEntry.RequestHeaders -join "`n" -replace '"', '""') + '"'),
                ('"' + ($RequestLogEntry.ResponseHeaders -replace '"', '""') + '"'),
                ('"' + ($RequestLogEntry.RequestBody -replace '"', '""') + '"'),
                ('"' + ($RequestLogEntry.ResponseBody -replace '"', '""') + '"'),
                $RequestLogEntry.ContentType,
                $RequestLogEntry.Time.ToString("0.000", [System.Globalization.CultureInfo]::InvariantCulture),
                $RequestLogEntry.ServerIP
            ) -join ";"
            
            # Append to CSV file
            Add-Content -Path $logFilePath -Value $csvLine -Encoding UTF8
        }
        catch {
            Write-Log "Waarschuwing: Kon niet schrijven naar request log: $_" -Level WARNING
        }
    }
}

# ============================================================================
# HTTP STATUS CODE INFORMATIE (uit OpenAPI Spec)
# ============================================================================

$script:HttpStatusInfo = @{
    200 = "OK - Request succesvol uitgevoerd"
    201 = "Created - Resource succesvol aangemaakt"
    204 = "No Content - Request succesvol, geen response body"
    400 = "Bad Request - Verplicht veld ontbreekt of heeft verkeerd formaat"
    401 = "Unauthorized - Authenticatie ontbreekt of is onjuist"
    403 = "Forbidden - Geen toegang tot deze resource"
    404 = "Not Found - Resource met opgegeven external ID bestaat niet"
    406 = "Not Acceptable - Gerefereerd object met external ID bestaat niet"
    409 = "Conflict - External ID is al toegewezen aan een andere resource"
    500 = "Internal Server Error - Interne serverfout"
    503 = "Service Unavailable - Service tijdelijk niet beschikbaar"
}

function Get-StatusCodeDescription {
    param([int]$StatusCode)
    
    if ($script:HttpStatusInfo.ContainsKey($StatusCode)) {
        return $script:HttpStatusInfo[$StatusCode]
    }
    return "HTTP $StatusCode"
}

# ============================================================================
# GENERIEKE REST METHOD WRAPPER MET LOGGING
# ============================================================================

function Invoke-RestMethodWithLogging {
    <#
    .SYNOPSIS
        Wrapper voor Invoke-RestMethod met automatische request logging
    #>
    param(
        [string]$Uri,
        [string]$Method = 'GET',
        [hashtable]$Headers,
        [object]$Body = $null,
        [string]$ContentType = 'application/json'
    )
    
    $startTime = Get-Date
    $requestHeaders = $Headers.GetEnumerator() | ForEach-Object { 
        $value = $_.Value
        # Mask token in headers for logging
        if ($_.Key -eq 'Authorization' -and $value -match 'Bearer (.+)') {
            $value = 'Bearer ***TOKEN***'
        }
        "$($_.Key): $value"
    }
    $requestBody = if ($Body) { $Body } else { "" }
    
    try {
        $params = @{
            Uri = $Uri
            Method = $Method
            Headers = $Headers
        }
        
        if ($Body) {
            $params['Body'] = $Body
        }
        
        if ($ContentType) {
            $params['ContentType'] = $ContentType
        }
        
        $response = Invoke-RestMethod @params
        $endTime = Get-Date
        $duration = ($endTime - $startTime).TotalSeconds
        
        $statusCode = 200
        $statusDescription = Get-StatusCodeDescription -StatusCode $statusCode
        
        # Log to CSV
        $logEntry = @{
            Timestamp = $startTime.ToString("yyyy-MM-dd HH:mm:ss.fff")
            Method = $Method
            URL = $Uri
            Status = $statusCode
            StatusText = $statusDescription
            RequestHeaders = $requestHeaders
            ResponseHeaders = ""
            RequestBody = $requestBody
            ResponseBody = if ($response) { ($response | ConvertTo-Json -Depth 3 -Compress) } else { "" }
            ContentType = $ContentType
            Time = $duration
            ServerIP = ""
        }
        Write-RequestLog -RequestLogEntry $logEntry
        
        return $response
    }
    catch {
        $endTime = Get-Date
        $duration = ($endTime - $startTime).TotalSeconds
        
        $statusCode = 500
        $statusDescription = "Error"
        $errorBody = ""
        
        if ($_.Exception.Response) {
            $statusCode = [int]$_.Exception.Response.StatusCode
            $statusDescription = Get-StatusCodeDescription -StatusCode $statusCode
            
            try {
                $reader = New-Object System.IO.StreamReader($_.Exception.Response.GetResponseStream())
                $errorBody = $reader.ReadToEnd()
                $reader.Close()
            }
            catch {
                $errorBody = $_.Exception.Message
            }
        }
        else {
            $errorBody = $_.Exception.Message
        }
        
        # Log error to CSV
        $logEntry = @{
            Timestamp = $startTime.ToString("yyyy-MM-dd HH:mm:ss.fff")
            Method = $Method
            URL = $Uri
            Status = $statusCode
            StatusText = $statusDescription
            RequestHeaders = $requestHeaders
            ResponseHeaders = ""
            RequestBody = $requestBody
            ResponseBody = $errorBody
            ContentType = $ContentType
            Time = $duration
            ServerIP = ""
        }
        Write-RequestLog -RequestLogEntry $logEntry
        
        throw
    }
}

# ============================================================================
# GLOBALE VARIABELEN
# ============================================================================

$script:currentSession = $null
$script:sessionHistory = @()
$script:baseUrl = if ($environment -eq 'sandbox') {
    "https://partners.sandbox.booqcloud.com"
} else {
    "https://partners.booqcloud.com"
}

$script:tokenUrl = if ($environment -eq 'production') {
	"https://partners.booqcloud.com/oauth2/token"
} else {
	"https://partners.sandbox.booqcloud.com/oauth2/token"
}

# API URLs worden dynamisch bepaald na token ophalen
$script:cmApiUrl = $null
$script:onboardingApiUrl = $null
$script:pimApiUrl = $null

$script:sessionFolder = Join-Path $script:startLocation "CM-Sessions"
if (-not (Test-Path $script:sessionFolder)) {
    New-Item -ItemType Directory -Path $script:sessionFolder -Force | Out-Null
}

# Cache voor referentie data
$script:cachedStores = $null
$script:cachedSalesPoints = $null
$script:cachedTurnoverGroups = $null
$script:cachedVatTariffs = $null
$script:cachedCustomers = $null
$script:cachedProducts = $null

# ============================================================================
# API URL INITIALISATIE
# ============================================================================

function Initialize-ApiUrls {
    Write-Log "API URLs initialiseren..." -level INFO
    
    $token = Get-ValidToken -ClientId $clientId -ClientSecret $clientSecret -TokenUrl $script:tokenUrl -ImpersonateClientId $impersonateClientId
    
    if (-not $token) {
        Write-Log "FOUT: Geen token verkregen" -level ERROR
        return $false
    }
    
    $hostUrl = Get-HostUrlFromToken -Token $token -Environment $environment
    
    if (-not $hostUrl) {
        Write-Log "FOUT: Geen host URL verkregen" -level ERROR
        return $false
    }
    
    $script:onboardingApiUrl = "https://$hostUrl/partner/onboarding/v1"
    $script:pimApiUrl = "https://$hostUrl/partner/pim/v1"
    $script:cmApiUrl = "https://$hostUrl/partner/cm/v1"
    
    Write-Log "Host URL: $hostUrl" -level SUCCESS
    Write-Log "Onboarding API: $($script:onboardingApiUrl)" -level INFO
    Write-Log "PIM API: $($script:pimApiUrl)" -level INFO
    Write-Log "CM API: $($script:cmApiUrl)" -level INFO
    
    if ($DebugMode) {
        Write-Log "=== API URL CONFIGURATIE ===" -level DEBUG
        Write-Log "Host URL uit token: $hostUrl" -level DEBUG
    }
    
    return $true
}

# ============================================================================
# SESSIE MANAGEMENT
# ============================================================================

function New-SessionObject {
    param(
        [string]$Name,
        [string]$Description
    )
    
    return [PSCustomObject]@{
        SessionId = [guid]::NewGuid().ToString()
        Name = $Name
        Description = $Description
        CreatedAt = Get-Date
        LastModified = Get-Date
        Environment = $environment
        ClientId = $clientId
        ImpersonateClientId = $impersonateClientId
        History = @()
        CreatedTimePeriods = @()
        CreatedAvailabilities = @()
        CreatedCustomers = @()
        CreatedPromotions = @()
        CreatedProducts = @()
        CreatedSalesPointGroups = @()
    }
}

function Save-Session {
    param([PSCustomObject]$Session)
    
    $Session.LastModified = Get-Date
    
    $tokenName = if ($impersonateClientId) { $impersonateClientId } else { $clientId }
    $safeTokenName = $tokenName -replace '[^a-zA-Z0-9]', '_'
    $safeSessionName = $Session.Name -replace '[^a-zA-Z0-9]', '_'
    
    $fileName = "cm-sessie_$($safeTokenName)_$($safeSessionName).xml"
    $filePath = Join-Path $script:sessionFolder $fileName
    
    try {
        $Session | Export-Clixml -Path $filePath -Depth 20 -Force
        Write-Log "Sessie opgeslagen: $filePath" -level SUCCESS
        
        # Extra validatie dat de file echt is opgeslagen
        if (Test-Path $filePath) {
            $fileInfo = Get-Item $filePath
            if ($DebugMode) {
                Write-Log "Sessie bestand grootte: $($fileInfo.Length) bytes" -level DEBUG
            }
        }
        
        return $true
    }
    catch {
        Write-Log "FOUT bij opslaan sessie: $_" -level ERROR
        return $false
    }
}

function Get-SavedSessions {
    $tokenName = if ($impersonateClientId) { $impersonateClientId } else { $clientId }
    $safeTokenName = $tokenName -replace '[^a-zA-Z0-9]', '_'
    
    $pattern = "cm-sessie_$($safeTokenName)_*.xml"
    $sessionFiles = Get-ChildItem -Path $script:sessionFolder -Filter $pattern -ErrorAction SilentlyContinue
    
    $sessions = @()
    foreach ($file in $sessionFiles) {
        try {
            $session = Import-Clixml -Path $file.FullName
            
            # Ensure alle properties bestaan (voor backward compatibility)
            if (-not ($session.PSObject.Properties.Name -contains 'CreatedProducts')) {
                $session | Add-Member -NotePropertyName "CreatedProducts" -NotePropertyValue @() -Force
            }
            if (-not ($session.PSObject.Properties.Name -contains 'CreatedSalesPointGroups')) {
                $session | Add-Member -NotePropertyName "CreatedSalesPointGroups" -NotePropertyValue @() -Force
            }
            if (-not ($session.PSObject.Properties.Name -contains 'History')) {
                $session | Add-Member -NotePropertyName "History" -NotePropertyValue @() -Force
            }
            if (-not ($session.PSObject.Properties.Name -contains 'CreatedTimePeriods')) {
                $session | Add-Member -NotePropertyName "CreatedTimePeriods" -NotePropertyValue @() -Force
            }
            if (-not ($session.PSObject.Properties.Name -contains 'CreatedAvailabilities')) {
                $session | Add-Member -NotePropertyName "CreatedAvailabilities" -NotePropertyValue @() -Force
            }
            if (-not ($session.PSObject.Properties.Name -contains 'CreatedCustomers')) {
                $session | Add-Member -NotePropertyName "CreatedCustomers" -NotePropertyValue @() -Force
            }
            if (-not ($session.PSObject.Properties.Name -contains 'CreatedPromotions')) {
                $session | Add-Member -NotePropertyName "CreatedPromotions" -NotePropertyValue @() -Force
            }
            
            # Originele structuur behouden voor compatibility
            $sessions += [PSCustomObject]@{
                FileName = $file.Name
                FilePath = $file.FullName
                Session = $session
            }
        }
        catch {
            Write-Log "WAARSCHUWING: Kan sessie niet laden: $($file.Name) - $_" -level WARNING
        }
    }
    
    return $sessions
}

function Show-SessionSelectionMenu {
    Clear-Host
    Write-Host "`n+==============================================================" -ForegroundColor Cyan
    Write-Host "|        BOOQ CM STAMDATA ONDERHOUD - SESSIE SELECTIE         |" -ForegroundColor Cyan
    Write-Host "+==============================================================+" -ForegroundColor Cyan
    
    try {
        $savedSessions = Get-SavedSessions
        $sessionCount = if ($savedSessions) { $savedSessions.Count } else { 0 }
        
        Write-Host "`n1. Nieuwe sessie aanmaken" -ForegroundColor Green
        
        if ($sessionCount -gt 0) {
            Write-Host "2. Bestaande sessie hervatten" -ForegroundColor Yellow
            Write-Host "3. Bestaande sessie kopieren met nieuwe naam" -ForegroundColor Yellow
        }
        
        Write-Host "Q. Afsluiten`n" -ForegroundColor Red
        
        $choice = Read-Host "Maak uw keuze"
        
        switch ($choice.ToUpper()) {
            "1" { 
                try {
                    $newSession = New-SessionFromUser
                    if ($newSession) {
                        Write-Log "Nieuwe sessie aangemaakt: $($newSession.Name)" -Level INFO
                    }
                    return $newSession
                }
                catch {
                    Write-Host "`n[FOUT] Fout bij aanmaken sessie: $_" -ForegroundColor Red
                    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
                    Write-Log "Fout bij aanmaken sessie: $_" -Level ERROR
                    Read-Host "`nDruk op Enter om opnieuw te proberen"
                    return Show-SessionSelectionMenu
                }
            }
            "2" { 
                if ($sessionCount -gt 0) {
                    try {
                        $session = Select-ExistingSession -Sessions $savedSessions -CopyMode $false
                        if ($session) {
                            Write-Log "Sessie hervat: $($session.Name)" -Level INFO
                        }
                        return $session
                    }
                    catch {
                        Write-Host "`n[FOUT] Fout bij laden sessie: $_" -ForegroundColor Red
                        Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Log "Fout bij laden sessie: $_" -Level ERROR
                        Read-Host "`nDruk op Enter om opnieuw te proberen"
                        return Show-SessionSelectionMenu
                    }
                }
                else {
                    Write-Host "`n[INFO] Geen opgeslagen sessies gevonden." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                    return Show-SessionSelectionMenu
                }
            }
            "3" { 
                if ($sessionCount -gt 0) {
                    try {
                        $session = Select-ExistingSession -Sessions $savedSessions -CopyMode $true
                        if ($session) {
                            Write-Log "Sessie gekopieerd: $($session.Name)" -Level INFO
                        }
                        return $session
                    }
                    catch {
                        Write-Host "`n[FOUT] Fout bij kopieren sessie: $_" -ForegroundColor Red
                        Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
                        Write-Log "Fout bij kopieren sessie: $_" -Level ERROR
                        Read-Host "`nDruk op Enter om opnieuw te proberen"
                        return Show-SessionSelectionMenu
                    }
                }
                else {
                    Write-Host "`n[INFO] Geen opgeslagen sessies gevonden." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                    return Show-SessionSelectionMenu
                }
            }
            "Q" { 
                Write-Host "`n[INFO] Script wordt afgesloten..." -ForegroundColor Yellow
                Write-Log "Script afgesloten door gebruiker in sessie selectie" -Level INFO
                exit 0
            }
            default { 
                Write-Host "`n[FOUT] Ongeldige keuze. Probeer opnieuw." -ForegroundColor Red
                Start-Sleep -Seconds 1
                return Show-SessionSelectionMenu
            }
        }
    }
    catch {
        Write-Host "`n[FOUT] Onverwachte fout in sessie selectie: $_" -ForegroundColor Red
        Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Onverwachte fout in Show-SessionSelectionMenu: $_" -Level ERROR
        Read-Host "`nDruk op Enter om opnieuw te proberen"
        return Show-SessionSelectionMenu
    }
}

function New-SessionFromUser {
    Clear-Host
    Write-Host "`n=== NIEUWE SESSIE AANMAKEN ===" -ForegroundColor Cyan
    
    $name = Read-Host "`nGeef een naam voor deze sessie"
    if ([string]::IsNullOrWhiteSpace($name)) {
        Write-Host "Sessienaam mag niet leeg zijn." -ForegroundColor Red
        Start-Sleep -Seconds 2
        return Show-SessionSelectionMenu
    }
    
    $description = Read-Host "Geef een beschrijving (optioneel)"
    
    $session = New-SessionObject -Name $name -Description $description
    [void](Save-Session -Session $session)
    
    Write-Host "`nSessie aangemaakt!" -ForegroundColor Green
    Start-Sleep -Seconds 1
    
    return $session
}

function Select-ExistingSession {
    param(
        [array]$Sessions,
        [bool]$CopyMode
    )
    
    Clear-Host
    $title = if ($CopyMode) { "SESSIE KOPIËREN" } else { "SESSIE SELECTEREN" }
    Write-Host "`n=== $title ===" -ForegroundColor Cyan
    
    Write-Host "`nBeschikbare sessies:`n" -ForegroundColor Yellow
    
    for ($i = 0; $i -lt $Sessions.Count; $i++) {
        $sess = $Sessions[$i].Session
        Write-Host "$($i + 1). $($sess.Name)" -ForegroundColor White
        Write-Host "   Beschrijving: $($sess.Description)" -ForegroundColor Gray
        Write-Host "   Laatst gewijzigd: $($sess.LastModified)" -ForegroundColor Gray
        Write-Host "   Aantal acties: $($sess.History.Count)" -ForegroundColor Gray
        Write-Host ""
    }
    
    Write-Host "B. Terug naar hoofdmenu`n" -ForegroundColor Red
    
    $choice = Read-Host "Selecteer een sessie"
    
    if ($choice -eq "B" -or $choice -eq "b") {
        return Show-SessionSelectionMenu
    }
    
    try {
        $index = [int]$choice - 1
        if ($index -ge 0 -and $index -lt $Sessions.Count) {
            $selectedSession = $Sessions[$index].Session
            
            if ($CopyMode) {
                $newName = Read-Host "`nGeef een nieuwe naam voor de kopie"
                if ([string]::IsNullOrWhiteSpace($newName)) {
                    Write-Host "Naam mag niet leeg zijn." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    return Select-ExistingSession -Sessions $Sessions -CopyMode $CopyMode
                }
                
                $copiedSession = $selectedSession.PSObject.Copy()
                $copiedSession.SessionId = [guid]::NewGuid().ToString()
                $copiedSession.Name = $newName
                $copiedSession.CreatedAt = Get-Date
                $copiedSession.LastModified = Get-Date
                
                [void](Save-Session -Session $copiedSession)
                Write-Host "`nSessie gekopieerd!" -ForegroundColor Green
                Start-Sleep -Seconds 1
                return $copiedSession
            }
            else {
                Write-Host "`nSessie geladen!" -ForegroundColor Green
                Start-Sleep -Seconds 1
                return $selectedSession
            }
        }
        else {
            Write-Host "Ongeldige selectie." -ForegroundColor Red
            Start-Sleep -Seconds 2
            return Select-ExistingSession -Sessions $Sessions -CopyMode $CopyMode
        }
    }
    catch {
        Write-Host "Ongeldige invoer." -ForegroundColor Red
        Start-Sleep -Seconds 2
        return Select-ExistingSession -Sessions $Sessions -CopyMode $CopyMode
    }
}

# ============================================================================
# API HELPER FUNCTIES
# ============================================================================

function Get-EnterpriseId {
    $token = Get-ValidToken -ClientId $clientId -ClientSecret $clientSecret -TokenUrl $script:tokenUrl -ImpersonateClientId $impersonateClientId -DebugMode $DebugMode
    return Get-EnterpriseIdFromToken -Token $token
}


function Get-HttpErrorResponseBody {
    param([Parameter(Mandatory)] $ErrorRecord)

    # PowerShell WebException path
    $ex = $ErrorRecord.Exception
    if ($ex -and $ex.Response) {
        try {
            $stream = $ex.Response.GetResponseStream()
            if ($stream) {
                $reader = New-Object System.IO.StreamReader($stream)
                $text = $reader.ReadToEnd()
                $reader.Close()
                return $text
            }
        } catch { }
    }

    # Invoke-RestMethod sometimes sets ErrorDetails.Message
    try {
        if ($ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) {
            return [string]$ErrorRecord.ErrorDetails.Message
        }
    } catch { }

    return $null
}

### 	ander techniek:	
### 	ander techniek:	function Invoke-CMApi {
### 	ander techniek:	    param(
### 	ander techniek:	        [string]$Endpoint,
### 	ander techniek:	        [string]$Method = "GET",
### 	ander techniek:	        [object]$Body = $null,
### 	ander techniek:	        [hashtable]$QueryParams = @{}
### 	ander techniek:	    )
### 	ander techniek:	
### 	ander techniek:	    $startTime = Get-Date
### 	ander techniek:	    $token = Get-ValidToken -ClientId $clientId -ClientSecret $clientSecret -TokenUrl $script:tokenUrl -ImpersonateClientId $impersonateClientId -DebugMode $DebugMode
### 	ander techniek:	    $enterpriseId = Get-EnterpriseId
### 	ander techniek:	
### 	ander techniek:	    $headers = @{
### 	ander techniek:	        "Authorization"      = "Bearer $token"
### 	ander techniek:	        "X-booq-enterpriseid"= $enterpriseId
### 	ander techniek:	        "Content-Type"       = "application/json"
### 	ander techniek:	        "Accept"             = "application/json"
### 	ander techniek:	    }
### 	ander techniek:	
### 	ander techniek:	    $uri = "$script:cmApiUrl/$Endpoint"
### 	ander techniek:	    if ($QueryParams.Count -gt 0) {
### 	ander techniek:	        $queryString = ($QueryParams.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '&'
### 	ander techniek:	        $uri += "?$queryString"
### 	ander techniek:	    }
### 	ander techniek:	
### 	ander techniek:	    $requestHeaders = $headers.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value -replace $token, '***TOKEN***')" }
### 	ander techniek:	    $requestBody = ""
### 	ander techniek:	
### 	ander techniek:	    Write-Log "API Call: $Method $uri" -level DEBUG
### 	ander techniek:	
### 	ander techniek:	    # Build request params
### 	ander techniek:	    $params = @{
### 	ander techniek:	        Uri         = $uri
### 	ander techniek:	        Method      = $Method
### 	ander techniek:	        Headers     = $headers
### 	ander techniek:	        ErrorAction = 'Stop'
### 	ander techniek:	    }
### 	ander techniek:	
### 	ander techniek:	    if ($Body -ne $null) {
### 	ander techniek:	        $requestBody = $Body | ConvertTo-Json -Depth 20 -Compress
### 	ander techniek:	        $params['Body'] = $requestBody
### 	ander techniek:	
### 	ander techniek:	        if ($DebugMode) {
### 	ander techniek:	            Write-Log "Request Body: $requestBody" -level DEBUG
### 	ander techniek:	        }
### 	ander techniek:	    }
### 	ander techniek:	
### 	ander techniek:	    try {
### 	ander techniek:	        # Use Invoke-WebRequest for consistent error body access
### 	ander techniek:	        $resp = Invoke-WebRequest @params
### 	ander techniek:	        $endTime = Get-Date
### 	ander techniek:	        $duration = ($endTime - $startTime).TotalSeconds
### 	ander techniek:	
### 	ander techniek:	        $statusCode = [int]$resp.StatusCode
### 	ander techniek:	        $statusDescription = Get-StatusCodeDescription -StatusCode $statusCode
### 	ander techniek:	
### 	ander techniek:	        Write-Log "[$statusCode] $Method $uri - $statusDescription" -level SUCCESS
### 	ander techniek:	
### 	ander techniek:	        $content = [string]$resp.Content
### 	ander techniek:	        $parsed = $null
### 	ander techniek:	        if (-not [string]::IsNullOrWhiteSpace($content)) {
### 	ander techniek:	            try { $parsed = $content | ConvertFrom-Json -ErrorAction Stop } catch { $parsed = $null }
### 	ander techniek:	        }
### 	ander techniek:	
### 	ander techniek:	        # CSV logging (success)
### 	ander techniek:	        $logEntry = @{
### 	ander techniek:	            Timestamp       = $startTime.ToString("yyyy-MM-dd HH:mm:ss.fff")
### 	ander techniek:	            Method          = $Method
### 	ander techniek:	            URL             = $uri
### 	ander techniek:	            Status          = $statusCode
### 	ander techniek:	            StatusText      = $statusDescription
### 	ander techniek:	            RequestHeaders  = $requestHeaders
### 	ander techniek:	            ResponseHeaders = ""   # optionally serialize $resp.Headers
### 	ander techniek:	            RequestBody     = $requestBody
### 	ander techniek:	            ResponseBody    = if ($parsed) { ($parsed | ConvertTo-Json -Depth 20 -Compress) } else { $content }
### 	ander techniek:	            ContentType     = "application/json"
### 	ander techniek:	            Time            = $duration
### 	ander techniek:	            ServerIP        = ""
### 	ander techniek:	        }
### 	ander techniek:	        Write-RequestLog -RequestLogEntry $logEntry
### 	ander techniek:	
### 	ander techniek:	        # Save original JSON if desired
### 	ander techniek:	        if ($saveOrgJson -and $content) {
### 	ander techniek:	            try {
### 	ander techniek:	                $apiCall = ($Endpoint + "_" + ($QueryParams.Keys -join "_")) -replace '[\\/:*?"<>|]', '_'
### 	ander techniek:	                $jsonFileName = "$enterpriseId-api-$apiCall.json"
### 	ander techniek:	                $jsonFilePath = Join-Path $script:jsonOutputFolder $jsonFileName
### 	ander techniek:	                $content | Out-File -FilePath $jsonFilePath -Encoding UTF8
### 	ander techniek:	                Write-Log "JSON response opgeslagen: $jsonFileName" -level DEBUG
### 	ander techniek:	            } catch {
### 	ander techniek:	                Write-Log "Waarschuwing: Kon JSON response niet opslaan: $_" -level WARNING
### 	ander techniek:	            }
### 	ander techniek:	        }
### 	ander techniek:	
### 	ander techniek:	        # Return parsed JSON when possible, else raw content
### 	ander techniek:	        if ($parsed) { return $parsed }
### 	ander techniek:	        if (-not [string]::IsNullOrWhiteSpace($content)) { return $content }
### 	ander techniek:	        return $null
### 	ander techniek:	    }
### 	ander techniek:		catch {
### 	ander techniek:			$endTime = Get-Date
### 	ander techniek:			$duration = ($endTime - $startTime).TotalSeconds
### 	ander techniek:		
### 	ander techniek:			$statusCode = 500
### 	ander techniek:			$statusDescription = "Error"
### 	ander techniek:			$errorBody = $null
### 	ander techniek:			$errorBody2 = $null
### 	ander techniek:			$errorBody3 = $null
### 	ander techniek:			$errorObj = $null
### 	ander techniek:			$responseHeadersText = ""
### 	ander techniek:		
### 	ander techniek:			if ($_.Exception.Response) {
### 	ander techniek:				try {
### 	ander techniek:					$statusCode = [int]$_.Exception.Response.StatusCode
### 	ander techniek:				} catch { }
### 	ander techniek:		
### 	ander techniek:				$statusDescription = Get-StatusCodeDescription -StatusCode $statusCode
### 	ander techniek:		
### 	ander techniek:				# Response headers (helpt zien of er überhaupt een body is)
### 	ander techniek:				try {
### 	ander techniek:					$hdrs = $_.Exception.Response.Headers
### 	ander techniek:					if ($hdrs) {
### 	ander techniek:						$responseHeadersText = ($hdrs.AllKeys | ForEach-Object { "${_}: $($hdrs[$_])" }) -join "`n"
### 	ander techniek:					}
### 	ander techniek:				} catch { }
### 	ander techniek:		
### 	ander techniek:				# Pad 1: WebException ResponseStream
### 	ander techniek:				try {
### 	ander techniek:					$stream = $_.Exception.Response.GetResponseStream()
### 	ander techniek:					if ($stream) {
### 	ander techniek:						$reader = New-Object System.IO.StreamReader($stream)
### 	ander techniek:						$errorBody = $reader.ReadToEnd()
### 	ander techniek:						$reader.Close()
### 	ander techniek:					}
### 	ander techniek:				} catch { }
### 	ander techniek:			}
### 	ander techniek:			else {
### 	ander techniek:				$statusDescription = Get-StatusCodeDescription -StatusCode $statusCode
### 	ander techniek:			}
### 	ander techniek:		
### 	ander techniek:			# Pad 2: PowerShell ErrorDetails.Message (hier zit vaak de body)
### 	ander techniek:			try {
### 	ander techniek:				if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
### 	ander techniek:					$errorBody2 = [string]$_.ErrorDetails.Message
### 	ander techniek:				}
### 	ander techniek:			} catch { }
### 	ander techniek:		
### 	ander techniek:			# Pad 3: Exception.Message (soms bevat die JSON bij sommige stacks)
### 	ander techniek:			try { $errorBody3 = [string]$_.Exception.Message } catch { }
### 	ander techniek:		
### 	ander techniek:			# Kies de beste body
### 	ander techniek:			$bodyCandidates = @($errorBody, $errorBody2) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
### 	ander techniek:			$finalBody = if ($bodyCandidates.Count -gt 0) { $bodyCandidates[0] } else { $null }
### 	ander techniek:		
### 	ander techniek:			# Probeer JSON parse
### 	ander techniek:			if (-not [string]::IsNullOrWhiteSpace($finalBody)) {
### 	ander techniek:				try { $errorObj = $finalBody | ConvertFrom-Json -ErrorAction Stop } catch { $errorObj = $null }
### 	ander techniek:			}
### 	ander techniek:		
### 	ander techniek:			Write-Log "[$statusCode] $Method $uri - $statusDescription" -level ERROR
### 	ander techniek:		
### 	ander techniek:			if (-not [string]::IsNullOrWhiteSpace($responseHeadersText) -and $DebugMode) {
### 	ander techniek:				Write-Log ("Response headers:`n{0}" -f $responseHeadersText) -level ERROR
### 	ander techniek:			}
### 	ander techniek:		
### 	ander techniek:			if ($errorObj -and ($errorObj.PSObject.Properties.Name -contains 'message')) {
### 	ander techniek:				Write-Log ("API message: {0}" -f $errorObj.message) -level ERROR
### 	ander techniek:				if ($errorObj.requestID)   { Write-Log ("requestID: {0}" -f $errorObj.requestID) -level ERROR }
### 	ander techniek:				if ($errorObj.traceID)     { Write-Log ("traceID: {0}" -f $errorObj.traceID) -level ERROR }
### 	ander techniek:				if ($errorObj.logStreamID) { Write-Log ("logStreamID: {0}" -f $errorObj.logStreamID) -level ERROR }
### 	ander techniek:		
### 	ander techniek:				if ($DebugMode) {
### 	ander techniek:					Write-Log ("Error details (raw): {0}" -f $finalBody) -level ERROR
### 	ander techniek:				}
### 	ander techniek:			}
### 	ander techniek:			else {
### 	ander techniek:				if ([string]::IsNullOrWhiteSpace($finalBody)) {
### 	ander techniek:					Write-Log "Error details: (leeg / geen response body)" -level ERROR
### 	ander techniek:					if ($DebugMode) {
### 	ander techniek:						if (-not [string]::IsNullOrWhiteSpace($errorBody3)) {
### 	ander techniek:							Write-Log ("Fallback Exception.Message: {0}" -f $errorBody3) -level ERROR
### 	ander techniek:						}
### 	ander techniek:					}
### 	ander techniek:				}
### 	ander techniek:				else {
### 	ander techniek:					Write-Log "Error details: $finalBody" -level ERROR
### 	ander techniek:				}
### 	ander techniek:			}
### 	ander techniek:		
### 	ander techniek:			# Log error to CSV (incl headers)
### 	ander techniek:			$logEntry = @{
### 	ander techniek:				Timestamp       = $startTime.ToString("yyyy-MM-dd HH:mm:ss.fff")
### 	ander techniek:				Method          = $Method
### 	ander techniek:				URL             = $uri
### 	ander techniek:				Status          = $statusCode
### 	ander techniek:				StatusText      = $statusDescription
### 	ander techniek:				RequestHeaders  = $requestHeaders
### 	ander techniek:				ResponseHeaders = $responseHeadersText
### 	ander techniek:				RequestBody     = $requestBody
### 	ander techniek:				ResponseBody    = if ($finalBody) { $finalBody } else { "" }
### 	ander techniek:				ContentType     = "application/json"
### 	ander techniek:				Time            = $duration
### 	ander techniek:				ServerIP        = ""
### 	ander techniek:			}
### 	ander techniek:			Write-RequestLog -RequestLogEntry $logEntry
### 	ander techniek:		
### 	ander techniek:			throw
### 	ander techniek:		}
### 	ander techniek:		
### 	ander techniek:	
### 	ander techniek:	}
### 	ander techniek:	
### 	ander techniek:	

	
	function Invoke-CMApi {
		param(
			[string]$Endpoint,
			[string]$Method = "GET",
			[object]$Body = $null,
			[hashtable]$QueryParams = @{}
		)
	
		$startTime = Get-Date
		$token = Get-ValidToken -ClientId $clientId -ClientSecret $clientSecret -TokenUrl $script:tokenUrl -ImpersonateClientId $impersonateClientId -DebugMode $DebugMode
		$enterpriseId = Get-EnterpriseId
	
		$headers = @{
			"Authorization"       = "Bearer $token"
			"X-booq-enterpriseid" = $enterpriseId
			"Accept"              = "application/json"
			"Content-Type"        = "application/json"
		}
	
		$uri = "$script:cmApiUrl/$Endpoint"
		if ($QueryParams.Count -gt 0) {
			$queryString = ($QueryParams.GetEnumerator() | ForEach-Object { "$($_.Key)=$($_.Value)" }) -join '&'
			$uri += "?$queryString"
		}
	
		$requestHeaders = $headers.GetEnumerator() | ForEach-Object { "$($_.Key): $($_.Value -replace $token, '***TOKEN***')" }
		$requestBody = ""
	
		Write-Log "API Call: $Method $uri" -level DEBUG
	
		$params = @{
			Uri         = $uri
			Method      = $Method
			Headers     = $headers
			ErrorAction = "Stop"
		}
	
		if ($Body -ne $null) {
			$requestBody = $Body | ConvertTo-Json -Depth 20 -Compress
			$params['Body'] = $requestBody
			if ($DebugMode) {
				Write-Log "Request Body: $requestBody" -level DEBUG
			}
		}
	
		try {
			$resp = Invoke-WebRequest @params
	
			$endTime = Get-Date
			$duration = ($endTime - $startTime).TotalSeconds
	
			$statusCode = [int]$resp.StatusCode
			$statusDescription = Get-StatusCodeDescription -StatusCode $statusCode
	
			Write-Log "[$statusCode] $Method $uri - $statusDescription" -level SUCCESS
	
			$content = [string]$resp.Content
	
			$logEntry = @{
				Timestamp       = $startTime.ToString("yyyy-MM-dd HH:mm:ss.fff")
				Method          = $Method
				URL             = $uri
				Status          = $statusCode
				StatusText      = $statusDescription
				RequestHeaders  = $requestHeaders
				ResponseHeaders = ""  # optioneel: $resp.Headers
				RequestBody     = $requestBody
				ResponseBody    = $content
				ContentType     = "application/json"
				Time            = $duration
				ServerIP        = ""
			}
			Write-RequestLog -RequestLogEntry $logEntry
	
			if ($saveOrgJson -and -not [string]::IsNullOrWhiteSpace($content)) {
				try {
					$apiCall = ($Endpoint + "_" + ($QueryParams.Keys -join "_")) -replace '[\\/:*?"<>|]', '_'
					$jsonFileName = "$enterpriseId-api-$apiCall.json"
					$jsonFilePath = Join-Path $script:jsonOutputFolder $jsonFileName
					$content | Out-File -FilePath $jsonFilePath -Encoding UTF8
					Write-Log "JSON response opgeslagen: $jsonFileName" -level DEBUG
				}
				catch {
					Write-Log "Waarschuwing: Kon JSON response niet opslaan: $_" -level WARNING
				}
			}
	
			if (-not [string]::IsNullOrWhiteSpace($content)) {
				try { return ($content | ConvertFrom-Json -ErrorAction Stop) } catch { return $content }
			}
			return $null
		}
		catch {
			$endTime = Get-Date
			$duration = ($endTime - $startTime).TotalSeconds
	
			$statusCode = 500
			$statusDescription = "Error"
			$finalBody = ""
			$errorObj = $null
			$responseHeadersText = ""
	
			if ($_.Exception -and $_.Exception.Response) {
				try { $statusCode = [int]$_.Exception.Response.StatusCode } catch { }
				$statusDescription = Get-StatusCodeDescription -StatusCode $statusCode
	
				# Headers alleen loggen in DebugMode (anders te veel ruis)
				if ($DebugMode) {
					try {
						$hdrs = $_.Exception.Response.Headers
						if ($hdrs) {
							$responseHeadersText = ($hdrs.AllKeys | ForEach-Object { "${_}: $($hdrs[$_])" }) -join "`n"
						}
					} catch { }
				}
	
				# Lees body volledig als bytes -> UTF8
				try {
					$stream = $_.Exception.Response.GetResponseStream()
					if ($stream) {
						$ms = New-Object System.IO.MemoryStream
						$stream.CopyTo($ms)
						$bytes = $ms.ToArray()
						$ms.Dispose()
						$finalBody = [System.Text.Encoding]::UTF8.GetString($bytes)
					}
				} catch { }
			}
			else {
				$statusDescription = Get-StatusCodeDescription -StatusCode $statusCode
			}
	
			# Fallback: ErrorDetails.Message
			try {
				if ([string]::IsNullOrWhiteSpace($finalBody) -and $_.ErrorDetails -and $_.ErrorDetails.Message) {
					$finalBody = [string]$_.ErrorDetails.Message
				}
			} catch { }
	
			# Parse JSON indien mogelijk
			if (-not [string]::IsNullOrWhiteSpace($finalBody)) {
				try { $errorObj = $finalBody | ConvertFrom-Json -ErrorAction Stop } catch { $errorObj = $null }
			}
	
			Write-Log "[$statusCode] $Method $uri - $statusDescription" -level ERROR
	
			if ($DebugMode -and -not [string]::IsNullOrWhiteSpace($responseHeadersText)) {
				Write-Log ("Response headers:`n{0}" -f $responseHeadersText) -level ERROR
			}
	
			# Kort en bruikbaar: API message + correlatie IDs
			if ($errorObj -and ($errorObj.PSObject.Properties.Name -contains 'message') -and -not [string]::IsNullOrWhiteSpace([string]$errorObj.message)) {
				Write-Log ("API message: {0}" -f ([string]$errorObj.message).Trim()) -level ERROR
				if ($errorObj.requestID)   { Write-Log ("requestID: {0}" -f $errorObj.requestID) -level ERROR }
				if ($errorObj.traceID)     { Write-Log ("traceID: {0}" -f $errorObj.traceID) -level ERROR }
				if ($errorObj.logStreamID) { Write-Log ("logStreamID: {0}" -f $errorObj.logStreamID) -level ERROR }
			}
			else {
				# Als het geen JSON is: toon 1 regel (trim) zodat het niet je scherm volspamt
				if ([string]::IsNullOrWhiteSpace($finalBody)) {
					Write-Log "Error details: (leeg / geen response body)" -level ERROR
				}
				else {
					$oneLine = ($finalBody -replace '\s+', ' ').Trim()
					if ($oneLine.Length -gt 400) { $oneLine = $oneLine.Substring(0, 400) + "..." }
					Write-Log ("Error details: {0}" -f $oneLine) -level ERROR
				}
			}
	
			# Raw body alleen in DebugMode
			if ($DebugMode -and -not [string]::IsNullOrWhiteSpace($finalBody)) {
				Write-Log ("Error details (raw): {0}" -f $finalBody) -level ERROR
			}
	
			# CSV: bewaar volledige body (handig voor later), ook als console kort is
			$logEntry = @{
				Timestamp       = $startTime.ToString("yyyy-MM-dd HH:mm:ss.fff")
				Method          = $Method
				URL             = $uri
				Status          = $statusCode
				StatusText      = $statusDescription
				RequestHeaders  = $requestHeaders
				ResponseHeaders = $responseHeadersText
				RequestBody     = $requestBody
				ResponseBody    = $finalBody
				ContentType     = "application/json"
				Time            = $duration
				ServerIP        = ""
			}
			Write-RequestLog -RequestLogEntry $logEntry
	
			throw
		}
	}
	


function Get-OnboardingData {
    param(
        [string]$Endpoint,
        [int]$Count = 500
    )
    
    if ($DebugMode) {
        Write-Log "=== GET ONBOARDING DATA ===" -level DEBUG
        Write-Log "Endpoint: $Endpoint" -level DEBUG
        Write-Log "Onboarding API URL: $($script:onboardingApiUrl)" -level DEBUG
    }
    
    if (-not $script:onboardingApiUrl) {
        Write-Log "FOUT: Onboarding API URL is niet geinitaliseerd!" -level ERROR
        Write-Log "Roep eerst Initialize-ApiUrls aan" -level ERROR
        return @()
    }
    
    $token = Get-ValidToken -ClientId $clientId -ClientSecret $clientSecret -TokenUrl $script:tokenUrl -ImpersonateClientId $impersonateClientId -DebugMode $DebugMode
    $enterpriseId = Get-EnterpriseId
    
    if ($DebugMode) {
        Write-Log "Enterprise ID: $enterpriseId" -level DEBUG
    }
    
    $headers = @{
        "Authorization" = "Bearer $token"
        "X-booq-enterpriseid" = $enterpriseId
    }
    
    $allItems = @()
    $from = $null
    $moreData = $true
    
    while ($moreData) {
        $uri = "$script:onboardingApiUrl/$Endpoint`?count=$Count"
        if ($from) {
            $uri += '&from=' + $from
        }
        
        if ($DebugMode) {
            Write-Log "Request URI: $uri" -level DEBUG
        }
        
        try {
            $response = Invoke-RestMethodWithLogging -Uri $uri -Method Get -Headers $headers
            
            if ($DebugMode) {
                Write-Log "Response properties: $($response.PSObject.Properties.Name -join ', ')" -level DEBUG
            }
            
            # Bepaal het juiste veld voor de items
            $itemsField = $null
            @('stores', 'salesPoints', 'turnoverGroups', 'vatTariffs', 'customers', 'paymentMethods', 'currencies') | ForEach-Object {
                if ($response.PSObject.Properties[$_]) {
                    $itemsField = $_
                }
            }
            
            if ($itemsField -and $response.$itemsField) {
                $allItems += $response.$itemsField
                if ($DebugMode) {
                    Write-Log "Gevonden: $itemsField met $($response.$itemsField.Count) items" -level DEBUG
                }
            }
            
            $moreData = $response.moreData -eq $true
            if ($moreData -and $response.marker) {
                $from = $response.marker
                if ($DebugMode) {
                    Write-Log "Meer data beschikbaar, marker: $from" -level DEBUG
                }
            }
        }
        catch {
            Write-Log "Fout bij ophalen $Endpoint`: $_" -level WARNING
            if ($DebugMode) {
                Write-Log "Fout details: $($_.Exception.Message)" -level DEBUG
                Write-Log "Stack trace: $($_.ScriptStackTrace)" -level DEBUG
            }
            break
        }
    }
    
    Write-Log "$Endpoint opgehaald: $($allItems.Count) items" -level SUCCESS
    return $allItems
}

function Get-PimProducts {
    param([int]$Count = 500)
    
    if ($DebugMode) {
        Write-Log "=== GET PIM PRODUCTS ===" -level DEBUG
        Write-Log "PIM API URL: $($script:pimApiUrl)" -level DEBUG
    }
    
    if (-not $script:pimApiUrl) {
        Write-Log "FOUT: PIM API URL is niet geinitaliseerd!" -level ERROR
        return @()
    }
    
    $token = Get-ValidToken -ClientId $clientId -ClientSecret $clientSecret -TokenUrl $script:tokenUrl -ImpersonateClientId $impersonateClientId -DebugMode $DebugMode
    $enterpriseId = Get-EnterpriseId
    
    if ($DebugMode) {
        Write-Log "Enterprise ID: $enterpriseId" -level DEBUG
    }
    
    $headers = @{
        "Authorization" = "Bearer $token"
        "X-booq-enterpriseid" = $enterpriseId
    }
    
    $allProducts = @()
    $from = $null
    $moreData = $true
    
    while ($moreData) {
        $uri = "$script:pimApiUrl/products?count=$Count"
        if ($from) {
            $uri += '&from=' + $from
        }
        
        if ($DebugMode) {
            Write-Log "Request URI: $uri" -level DEBUG
        }
        
        try {
            $response = Invoke-RestMethodWithLogging -Uri $uri -Method Get -Headers $headers
            
            if ($response.products) {
                $allProducts += $response.products
                if ($DebugMode) {
                    Write-Log "Batch opgehaald: $($response.products.Count) producten" -level DEBUG
                }
            }
            
            $moreData = $response.moreData -eq $true
            if ($moreData -and $response.marker) {
                $from = $response.marker
            }
        }
        catch {
            Write-Log "Fout bij ophalen products: $_" -level WARNING
            if ($DebugMode) {
                Write-Log "Fout details: $($_.Exception.Message)" -level DEBUG
            }
            break
        }
    }
    
    Write-Log "Products opgehaald: $($allProducts.Count) items" -level SUCCESS
    return $allProducts
}

function Show-SalesPointMultiSelect {
    param(
        [Parameter(Mandatory=$true)]
        [array]$SalesPoints
    )
    
    # FIX v1.3: Forms assembly loading met error handling
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    }
    catch {
        Write-Log "FOUT: Windows Forms kon niet worden geladen: $_" -level ERROR
        Write-Host "[FOUT] GUI kan niet worden gestart. Forms assembly niet beschikbaar." -ForegroundColor Red
        throw "Windows Forms not available"
    }
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Selecteer SalesPoints voor SalesPointGroup"
    $form.Size = New-Object System.Drawing.Size(600, 500)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(580, 20)
    $label.Text = "Selecteer salespoints (niveau 0 en 1):"
    $form.Controls.Add($label)
    
    $checkedListBox = New-Object System.Windows.Forms.CheckedListBox
    $checkedListBox.Location = New-Object System.Drawing.Point(10, 35)
    $checkedListBox.Size = New-Object System.Drawing.Size(560, 360)
    $checkedListBox.CheckOnClick = $true
    
    # Hashtable om de mapping tussen index en SalesPoint ID bij te houden
    $indexToIdMapping = @{}
    
    $level0and1 = $SalesPoints | Where-Object { $_.level -le 1 }
    
    $itemIndex = 0
    foreach ($sp in $level0and1) {
        $indent = "  " * $sp.level
        $displayText = "$indent$($sp.name) [$($sp.id)]"
        if ($sp.level -eq 0) {
            $displayText = "$($sp.name) [Store-level SP: $($sp.id)]"
        }
        
        # Voeg item toe en bewaar de mapping
        $checkedListBox.Items.Add($displayText) | Out-Null
        $indexToIdMapping[$itemIndex] = $sp.id
        $itemIndex++
    }
    
    $form.Controls.Add($checkedListBox)
    
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(400, 410)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton
    
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(490, 410)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.Text = "Annuleren"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton
    
    # FIX v1.3: ShowDialog met error handling
    try {
        $form.TopMost = $true
        $result = $form.ShowDialog()
    }
    catch {
        Write-Log "FOUT: ShowDialog gefaald: $_" -level ERROR
        $result = [System.Windows.Forms.DialogResult]::Cancel
    }
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedIds = @()
        foreach ($index in $checkedListBox.CheckedIndices) {
            # Haal de SalesPoint ID op uit de mapping
            $spId = $indexToIdMapping[$index]
            if ($spId) {
                $selectedIds += $spId
            }
        }
        
        if ($DebugMode) {
            Write-Log "=== SALESPOINT MULTI-SELECT RESULT ===" -level DEBUG
            Write-Log "Aantal geselecteerd: $($selectedIds.Count)" -level DEBUG
            Write-Log "Geselecteerde IDs: $($selectedIds -join ', ')" -level DEBUG
        }
        
        return $selectedIds
    }
    
    return $null
}


function Show-CustomerMultiSelect {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Customers,
        [array]$PreSelectedIds = @()
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    }
    catch {
        Write-Log "FOUT: Windows Forms kon niet worden geladen: $_" -level ERROR
        Write-Host "[FOUT] GUI kan niet worden gestart. Forms assembly niet beschikbaar." -ForegroundColor Red
        throw "Windows Forms not available"
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Selecteer Customers"
    $form.Size = New-Object System.Drawing.Size(760, 640)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(730, 40)
    $label.Text = "Selecteer klanten:`nZoeken filtert de lijst maar behoudt geselecteerde klanten."
    $form.Controls.Add($label)

    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Location = New-Object System.Drawing.Point(10, 55)
    $searchLabel.Size = New-Object System.Drawing.Size(100, 20)
    $searchLabel.Text = "Zoeken:"
    $form.Controls.Add($searchLabel)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(110, 52)
    $searchBox.Size = New-Object System.Drawing.Size(630, 25)
    $form.Controls.Add($searchBox)

    $checkedListBox = New-Object System.Windows.Forms.CheckedListBox
    $checkedListBox.Location = New-Object System.Drawing.Point(10, 85)
    $checkedListBox.Size = New-Object System.Drawing.Size(730, 450)
    $checkedListBox.CheckOnClick = $true
    $checkedListBox.Sorted = $false
    $form.Controls.Add($checkedListBox)

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Location = New-Object System.Drawing.Point(10, 540)
    $infoLabel.Size = New-Object System.Drawing.Size(730, 20)
    $form.Controls.Add($infoLabel)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(550, 570)
    $okButton.Size = New-Object System.Drawing.Size(90, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(650, 570)
    $cancelButton.Size = New-Object System.Drawing.Size(90, 30)
    $cancelButton.Text = "Annuleren"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton

    # Centrale selectie (id -> true)
    $selected = @{}
    foreach ($id in @($PreSelectedIds)) {
        $sid = ([string]$id).Trim()
        if ($sid) { $selected[$sid] = $true }
    }

    # View mapping: index -> id
    $indexToId = @{}
    $suppress = $false

    # Normaliseer customer list
    $all = @()
    foreach ($c in $Customers) {
        $id = ([string]$c.id).Trim()
        if (-not $id) { continue }
        $name = [string]$c.name
        $display = if ($name) { "$name (ID: $id)" } else { "ID: $id" }
        $all += [PSCustomObject]@{ Id = $id; Name = $name; Display = $display }
    }

    function Update-Info {
        $infoLabel.Text = "Zichtbaar: $($checkedListBox.Items.Count) | Geselecteerd totaal: $($selected.Keys.Count)"
    }

    function Render-List([string]$filter) {
        $script:suppress = $true
        try {
            $checkedListBox.Items.Clear()
            $indexToId.Clear()

            $ft = if ($filter) { $filter.ToLower() } else { "" }

            $i = 0
            foreach ($c in $all) {
                $match =
                    [string]::IsNullOrWhiteSpace($ft) -or
                    ($c.Display.ToLower().Contains($ft)) -or
                    ($c.Id.ToLower().Contains($ft))

                if (-not $match) { continue }

                [void]$checkedListBox.Items.Add($c.Display)
                $indexToId[$i] = $c.Id

                if ($selected.ContainsKey($c.Id)) {
                    $checkedListBox.SetItemChecked($i, $true)
                }
                $i++
            }

            Update-Info
        }
        finally {
            $script:suppress = $false
        }
    }

    $checkedListBox.Add_ItemCheck({
        param($sender, $e)

        if ($script:suppress) { return }

        $idx = [int]$e.Index
        if (-not $indexToId.ContainsKey($idx)) { return }
        $id = $indexToId[$idx]

        if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
            $selected[$id] = $true
        } else {
            if ($selected.ContainsKey($id)) { $selected.Remove($id) }
        }

        $form.BeginInvoke([Action]{ Update-Info }) | Out-Null
    })

    $searchBox.Add_TextChanged({
        Render-List $searchBox.Text
    })

    Render-List ""

    $form.TopMost = $true
    $result = $form.ShowDialog()
    $form.Dispose()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedIds = @($selected.Keys | Sort-Object)
        if ($DebugMode) {
            Write-Log "=== CUSTOMER MULTI-SELECT RESULT ===" -level DEBUG
            Write-Log "Aantal geselecteerd: $($selectedIds.Count)" -level DEBUG
            Write-Log "Geselecteerde IDs: $($selectedIds -join ', ')" -level DEBUG
        }
        return $selectedIds
    }

    return $null
}



function Show-ProductMultiSelect {
    param(
        [Parameter(Mandatory=$true)]
        [array]$Products,
        [array]$PreSelectedIds = @()
    )

    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    }
    catch {
        Write-Log "FOUT: Windows Forms kon niet worden geladen: $_" -level ERROR
        Write-Host "[FOUT] GUI kan niet worden gestart. Forms assembly niet beschikbaar." -ForegroundColor Red
        throw "Windows Forms not available"
    }

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Selecteer Producten"
    $form.Size = New-Object System.Drawing.Size(820, 720)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    $form.MinimizeBox = $false

    # Centrale selectie (id -> $true). Blijft bestaan tijdens filteren.
    $selected = @{}
    foreach ($id in @($PreSelectedIds)) {
        $sid = ([string]$id).Trim()
        if ($sid) { $selected[$sid] = $true }
    }

    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(780, 40)
    $label.Text = "Selecteer producten voor de promotie:`nZoeken/filteren behoudt geselecteerde producten."
    $form.Controls.Add($label)

    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Location = New-Object System.Drawing.Point(10, 55)
    $searchLabel.Size = New-Object System.Drawing.Size(100, 20)
    $searchLabel.Text = "Zoeken:"
    $form.Controls.Add($searchLabel)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(110, 52)
    $searchBox.Size = New-Object System.Drawing.Size(660, 25)
    $form.Controls.Add($searchBox)

    $filterLabel = New-Object System.Windows.Forms.Label
    $filterLabel.Location = New-Object System.Drawing.Point(10, 85)
    $filterLabel.Size = New-Object System.Drawing.Size(100, 20)
    $filterLabel.Text = "Zoek in:"
    $form.Controls.Add($filterLabel)

    $filterCombo = New-Object System.Windows.Forms.ComboBox
    $filterCombo.Location = New-Object System.Drawing.Point(110, 82)
    $filterCombo.Size = New-Object System.Drawing.Size(200, 25)
    $filterCombo.DropDownStyle = "DropDownList"
    [void]$filterCombo.Items.Add("Naam en ID")
    [void]$filterCombo.Items.Add("Alleen Naam")
    [void]$filterCombo.Items.Add("Alleen ID")
    $filterCombo.SelectedIndex = 0
    $form.Controls.Add($filterCombo)

    $selectAllBtn = New-Object System.Windows.Forms.Button
    $selectAllBtn.Location = New-Object System.Drawing.Point(320, 81)
    $selectAllBtn.Size = New-Object System.Drawing.Size(110, 26)
    $selectAllBtn.Text = "Alles Selecteren"
    $form.Controls.Add($selectAllBtn)

    $deselectAllBtn = New-Object System.Windows.Forms.Button
    $deselectAllBtn.Location = New-Object System.Drawing.Point(440, 81)
    $deselectAllBtn.Size = New-Object System.Drawing.Size(130, 26)
    $deselectAllBtn.Text = "Alles Deselecteren"
    $form.Controls.Add($deselectAllBtn)

    $checkedListBox = New-Object System.Windows.Forms.CheckedListBox
    $checkedListBox.Location = New-Object System.Drawing.Point(10, 115)
    $checkedListBox.Size = New-Object System.Drawing.Size(780, 520)
    $checkedListBox.CheckOnClick = $true
    $checkedListBox.Sorted = $false
    $form.Controls.Add($checkedListBox)

    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Location = New-Object System.Drawing.Point(10, 645)
    $countLabel.Size = New-Object System.Drawing.Size(780, 20)
    $form.Controls.Add($countLabel)

    # view mapping: index -> productId
    $indexToId = @{}
    $suppress = $false

    # Normaliseer bronlijst één keer
    $all = @()
    foreach ($p in $Products) {
        $id = ([string]$p.id).Trim()
        if (-not $id) { continue }
        $name = [string]$p.name
        $display = if ($name) { "$name (ID: $id)" } else { "ID: $id" }
        $all += [PSCustomObject]@{ Id = $id; Name = $name; Display = $display }
    }

    function Update-CountLabel {
        $countLabel.Text = "Zichtbaar: $($checkedListBox.Items.Count) | Geselecteerd totaal: $($selected.Keys.Count)"
    }

    function Render-List {
        param([string]$searchText, [int]$filterType)

        $script:suppress = $true
        try {
            $checkedListBox.Items.Clear()
            $indexToId.Clear()

            $st = if ($searchText) { $searchText.ToLower() } else { "" }

            $i = 0
            foreach ($p in $all) {
                $match = $true
                if (-not [string]::IsNullOrWhiteSpace($st)) {
                    switch ($filterType) {
                        0 { $match = ($p.Name.ToLower().Contains($st) -or $p.Id.ToLower().Contains($st)) }
                        1 { $match = ($p.Name.ToLower().Contains($st)) }
                        2 { $match = ($p.Id.ToLower().Contains($st)) }
                    }
                }

                if (-not $match) { continue }

                [void]$checkedListBox.Items.Add($p.Display)
                $indexToId[$i] = $p.Id

                if ($selected.ContainsKey($p.Id)) {
                    $checkedListBox.SetItemChecked($i, $true)
                }
                $i++
            }

            Update-CountLabel
        }
        finally {
            $script:suppress = $false
        }
    }

    # Belangrijk: gebruik ItemCheckEventArgs.Index, niet SelectedIndex
    $checkedListBox.Add_ItemCheck({
        param($sender, $e)

        if ($script:suppress) { return }

        $idx = [int]$e.Index
        if (-not $indexToId.ContainsKey($idx)) { return }
        $id = $indexToId[$idx]

        if ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) {
            $selected[$id] = $true
        } else {
            if ($selected.ContainsKey($id)) { $selected.Remove($id) }
        }

        $form.BeginInvoke([Action]{
            Update-CountLabel
        }) | Out-Null
    })

    $searchBox.Add_TextChanged({
        Render-List -searchText $searchBox.Text -filterType $filterCombo.SelectedIndex
    })
    $filterCombo.Add_SelectedIndexChanged({
        Render-List -searchText $searchBox.Text -filterType $filterCombo.SelectedIndex
    })

    $selectAllBtn.Add_Click({
        foreach ($p in $all) { $selected[$p.Id] = $true }
        Render-List -searchText $searchBox.Text -filterType $filterCombo.SelectedIndex
    })

    $deselectAllBtn.Add_Click({
        $selected.Clear()
        Render-List -searchText $searchBox.Text -filterType $filterCombo.SelectedIndex
    })

    # Buttons
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(600, 670)
    $okButton.Size = New-Object System.Drawing.Size(90, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($okButton)
    $form.AcceptButton = $okButton

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(700, 670)
    $cancelButton.Size = New-Object System.Drawing.Size(90, 30)
    $cancelButton.Text = "Annuleren"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancelButton)
    $form.CancelButton = $cancelButton

    # initial render
    Render-List -searchText "" -filterType 0

    $form.TopMost = $true
    $result = $form.ShowDialog()
    $form.Dispose()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedIds = @($selected.Keys | Sort-Object)
        if ($DebugMode) {
            Write-Log "=== PRODUCT MULTI-SELECT RESULT ===" -level DEBUG
            Write-Log "Aantal geselecteerd: $($selectedIds.Count)" -level DEBUG
            Write-Log "Geselecteerde IDs: $($selectedIds -join ', ')" -level DEBUG
        }
        return $selectedIds
    }

    return $null
}

function New-SalesPointGroup {
    param(
        [string]$GroupId,
        [string]$Name,
        [array]$SalesPointIds
    )
    
    Write-Log "SalesPointGroup aanmaken: $GroupId" -level INFO
    
    if ($DebugMode) {
        Write-Log "=== NEW SALESPOINTGROUP ===" -level DEBUG
        Write-Log "Group ID: $GroupId" -level DEBUG
        Write-Log "Name: $Name" -level DEBUG
        Write-Log "SalesPoints: $($SalesPointIds -join ', ')" -level DEBUG
    }
    
    $body = @{
        name = $Name
        salesPoints = $SalesPointIds
    }
    
    try {
        $response = Invoke-CMApi -Endpoint "salespointgroups/$GroupId" -Method "PUT" -Body $body
        Write-Log "SalesPointGroup aangemaakt: $GroupId" -level SUCCESS
        return $response
    }
    catch {
        Write-Log "Fout bij aanmaken SalesPointGroup: $_" -level ERROR
        return $null
    }
}

function Select-SalesPointGroup {
    param([PSCustomObject]$Session)
    
    Clear-Host
    Write-Host "`n=== SALESPOINTGROUP SELECTEREN ===" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. Nieuwe SalesPointGroup aanmaken" -ForegroundColor Green
    Write-Host "2. Bestaande selecteren uit sessie" -ForegroundColor Yellow
    Write-Host "3. Handmatig ID invoeren" -ForegroundColor Gray
    
    $choice = Read-Host "`nKeuze [1-3]"
    
    switch ($choice) {
        "1" {
            $groupId = Read-InputWithValidation -Prompt "Geef unieke ID voor SalesPointGroup" -Type "String" -Mandatory $true
            $name = Read-InputWithValidation -Prompt "Naam van de SalesPointGroup" -Type "String" -Mandatory $true
            
            Write-Host "`nSalesPoints ophalen..." -ForegroundColor Yellow
            $salesPoints = Get-CachedSalesPoints
            
            if (-not $salesPoints -or $salesPoints.Count -eq 0) {
                Write-Host "Geen salespoints beschikbaar!" -ForegroundColor Red
                Read-Host "`nDruk op Enter"
                return Select-SalesPointGroup -Session $Session
            }
            
            Write-Host "Multi-select venster wordt geopend..." -ForegroundColor Yellow
            $selectedIds = Show-SalesPointMultiSelect -SalesPoints $salesPoints
            
            if (-not $selectedIds -or $selectedIds.Count -eq 0) {
                Write-Host "Geen salespoints geselecteerd!" -ForegroundColor Red
                Read-Host "`nDruk op Enter"
                return Select-SalesPointGroup -Session $Session
            }
            
            Write-Host "`nGeselecteerd: $($selectedIds.Count) salespoints" -ForegroundColor Green
            Write-Host "SalesPointGroup wordt aangemaakt..." -ForegroundColor Yellow
            
            $result = New-SalesPointGroup -GroupId $groupId -Name $name -SalesPointIds $selectedIds
            
            if ($result) {
                $spgInfo = [PSCustomObject]@{
                    id = $groupId
                    name = $name
                    salesPointIds = $selectedIds
                }
                
                if (-not $Session.CreatedSalesPointGroups) {
                    $Session | Add-Member -NotePropertyName "CreatedSalesPointGroups" -NotePropertyValue @() -Force
                }
                $Session.CreatedSalesPointGroups += $spgInfo
                
                Write-Host "SalesPointGroup aangemaakt!" -ForegroundColor Green
                Read-Host "`nDruk op Enter"
                return [string]$groupId
            }
            else {
                Write-Host "Fout bij aanmaken!" -ForegroundColor Red
                Read-Host "`nDruk op Enter"
                return $null
            }
        }
        "2" {
            if ($Session.CreatedSalesPointGroups -and $Session.CreatedSalesPointGroups.Count -gt 0) {
                $selected = Show-SelectionMenu -Title "Selecteer SalesPointGroup" -Items $Session.CreatedSalesPointGroups -DisplayProperty "name" -IdProperty "id" -AllowCancel $true
                if ($selected) {
                    return [string]$selected.id
                }
            }
            else {
                Write-Host "Geen SalesPointGroups in sessie!" -ForegroundColor Red
                Read-Host "`nDruk op Enter"
            }
            return Select-SalesPointGroup -Session $Session
        }
        "3" {
            $groupId = Read-InputWithValidation -Prompt "Geef SalesPointGroup ID" -Type "String" -Mandatory $true
            return [string]$groupId
        }
        default {
            return Select-SalesPointGroup -Session $Session
        }
    }
}

function Get-SalesPointsFromStores {
    Write-Log "SalesPoints ophalen via stores..." -level INFO
    
    if (-not $script:onboardingApiUrl) {
        Write-Log "FOUT: Onboarding API URL is niet geinitaliseerd!" -level ERROR
        return @()
    }
    
    $stores = Get-OnboardingData -Endpoint "stores"
    
    if (-not $stores -or $stores.Count -eq 0) {
        Write-Log "Geen stores gevonden" -level WARNING
        return @()
    }
    
    Write-Log "Stores gevonden: $($stores.Count)" -level INFO
    
    $token = Get-ValidToken -ClientId $clientId -ClientSecret $clientSecret -TokenUrl $script:tokenUrl -ImpersonateClientId $impersonateClientId
    $enterpriseId = Get-EnterpriseId
    
    $headers = @{
        "Authorization" = "Bearer $token"
        "X-booq-enterpriseid" = $enterpriseId
    }
    
    $allSalesPoints = @()
    
    foreach ($store in $stores) {
        Write-Log "  Salespoints ophalen voor store: $($store.id) - $($store.name)" -level INFO
        
        $spEndpoint = "stores/$($store.id)/salespoints"
        $uri = "$script:onboardingApiUrl/$spEndpoint"
        
        if ($DebugMode) {
            Write-Log "=== SALESPOINTS REQUEST ===" -level DEBUG
            Write-Log "Store ID: $($store.id)" -level DEBUG
            Write-Log "URI: $uri" -level DEBUG
        }
        
        try {
            $response = Invoke-RestMethodWithLogging -Uri $uri -Method Get -Headers $headers
            
            if ($DebugMode) {
                Write-Log "Response properties: $($response.PSObject.Properties.Name -join ', ')" -level DEBUG
                if ($response.storeSalesPointId) {
                    Write-Log "Store SalesPoint ID uit API: $($response.storeSalesPointId)" -level DEBUG
                }
            }
            
            # Gebruik de storeSalesPointId uit de API response
            if ($response -and $response.storeSalesPointId) {
                $storeSalesPoint = [PSCustomObject]@{
                    id = $response.storeSalesPointId
                    name = $store.name
                    parentId = $null
                    storeName = $store.name
                    storeId = $store.id
                    level = 0
                }
                $allSalesPoints += $storeSalesPoint
                
                Write-Log "    Store-level SalesPoint ID: $($response.storeSalesPointId)" -level INFO
                
                if ($response.salesPoints) {
                    foreach ($sp in $response.salesPoints) {
                        $sp | Add-Member -NotePropertyName "parentId" -NotePropertyValue $response.storeSalesPointId -Force
                        $sp | Add-Member -NotePropertyName "storeName" -NotePropertyValue $store.name -Force
                        $sp | Add-Member -NotePropertyName "storeId" -NotePropertyValue $store.id -Force
                        $sp | Add-Member -NotePropertyName "level" -NotePropertyValue 1 -Force
                        
                        $allSalesPoints += $sp
                    }
                    
                    Write-Log "  Store $($store.id): $($response.salesPoints.Count) child salespoints" -level SUCCESS
                }
            }
            else {
                Write-Log "  Store $($store.id): geen storeSalesPointId in response" -level WARNING
            }
        }
        catch {
            Write-Log "  Fout bij store $($store.id): $_" -level WARNING
        }
    }
    
    Write-Log "Totaal SalesPoints: $($allSalesPoints.Count)" -level SUCCESS
    return $allSalesPoints
}


function New-SalesPointGroupCore {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Session,

        [Parameter(Mandatory)]
        [string]$GroupId,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [string[]]$SalesPointIds
    )

    $GroupId = $GroupId.Trim()
    $Name = $Name.Trim()
    $SalesPointIds = @($SalesPointIds | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })

    # API call (zoals bij jou OK ging)
    $body = @{
        name        = $Name
        salesPoints = $SalesPointIds
    }
    $null = Invoke-CMApi -Endpoint "salespointgroups/$GroupId" -Method "PUT" -Body $body

    # Zorg dat property bestaat
    if (-not $Session.PSObject.Properties.Match('CreatedSalesPointGroups')) {
        $Session | Add-Member -NotePropertyName CreatedSalesPointGroups -NotePropertyValue @() -Force
    }
    if ($null -eq $Session.CreatedSalesPointGroups) { $Session.CreatedSalesPointGroups = @() }

    Write-Host "DEBUG SPG: vóór add, count = $($Session.CreatedSalesPointGroups.Count)"

    $spgInfo = [PSCustomObject]@{
        id           = $GroupId
        name         = $Name
        salesPointIds = $SalesPointIds
    }

    # Append (of update existing)
    $found = $false
    for ($i=0; $i -lt $Session.CreatedSalesPointGroups.Count; $i++) {
        $item = $Session.CreatedSalesPointGroups[$i]
        $itemId = if ($item -is [string]) { [string]$item } elseif ($item.PSObject.Properties['id']) { [string]$item.id } else { "" }
        if ($itemId -eq $GroupId) {
            $Session.CreatedSalesPointGroups[$i] = $spgInfo
            $found = $true
            break
        }
    }
    if (-not $found) {
        $Session.CreatedSalesPointGroups += $spgInfo
    }

    Write-Host "DEBUG SPG: na add, count = $($Session.CreatedSalesPointGroups.Count)"
    Write-Host "DEBUG SPG: laatste id = $((@($Session.CreatedSalesPointGroups)[-1]).id)"

    [void](Save-Session -Session $Session)
    return $GroupId
}

function Show-SalesPointGroupSelectorGUI {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Session,
        [System.Windows.Forms.Form]$OwnerForm
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "SalesPointGroup selecteren / aanmaken"
    $dlg.StartPosition = if ($OwnerForm) { "CenterParent" } else { "CenterScreen" }
    $dlg.Size = New-Object System.Drawing.Size(860, 520)
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false

    $dlg.Tag = @() # selected salesPointIds
    $selectedId = $null

    # Left
    $grpList = New-Object System.Windows.Forms.GroupBox
    $grpList.Location = New-Object System.Drawing.Point(16, 16)
    $grpList.Size = New-Object System.Drawing.Size(400, 420)
    $grpList.Text = "Bestaande SalesPointGroups (sessie)"
    $dlg.Controls.Add($grpList)

    $lst = New-Object System.Windows.Forms.ListBox
    $lst.Location = New-Object System.Drawing.Point(12, 24)
    $lst.Size = New-Object System.Drawing.Size(376, 330)
    $lst.DisplayMember = "name"
    $grpList.Controls.Add($lst)

    $lblCount = New-Object System.Windows.Forms.Label
    $lblCount.Location = New-Object System.Drawing.Point(12, 362)
    $lblCount.Size = New-Object System.Drawing.Size(376, 20)
    $grpList.Controls.Add($lblCount)

    $btnUse = New-Object System.Windows.Forms.Button
    $btnUse.Location = New-Object System.Drawing.Point(12, 386)
    $btnUse.Size = New-Object System.Drawing.Size(376, 28)
    $btnUse.Text = "Gebruik geselecteerde"
    $grpList.Controls.Add($btnUse)

    function Refresh-List {
        $lst.Items.Clear()
        $items = @()
        if ($Session.CreatedSalesPointGroups) { $items = @($Session.CreatedSalesPointGroups) }

        foreach ($spg in $items) {
            if ($spg -is [string]) {
                [void]$lst.Items.Add([PSCustomObject]@{ id = $spg; name = $spg; salesPointIds = @() })
            } else {
                [void]$lst.Items.Add($spg)
            }
        }

        $cnt = if ($Session.CreatedSalesPointGroups) { $Session.CreatedSalesPointGroups.Count } else { 0 }
        $lblCount.Text = "In sessie: $cnt"
    }

    # Right
    $grpNew = New-Object System.Windows.Forms.GroupBox
    $grpNew.Location = New-Object System.Drawing.Point(440, 16)
    $grpNew.Size = New-Object System.Drawing.Size(390, 420)
    $grpNew.Text = "Nieuwe SalesPointGroup aanmaken"
    $dlg.Controls.Add($grpNew)

    $lblId = New-Object System.Windows.Forms.Label
    $lblId.Location = New-Object System.Drawing.Point(12, 28)
    $lblId.Size = New-Object System.Drawing.Size(100, 20)
    $lblId.Text = "GroupId:"
    $grpNew.Controls.Add($lblId)

    $txtId = New-Object System.Windows.Forms.TextBox
    $txtId.Location = New-Object System.Drawing.Point(120, 26)
    $txtId.Size = New-Object System.Drawing.Size(250, 20)
    $grpNew.Controls.Add($txtId)

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Location = New-Object System.Drawing.Point(12, 62)
    $lblName.Size = New-Object System.Drawing.Size(100, 20)
    $lblName.Text = "Naam:"
    $grpNew.Controls.Add($lblName)

    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(120, 60)
    $txtName.Size = New-Object System.Drawing.Size(250, 20)
    $grpNew.Controls.Add($txtName)

    $btnPickSalesPoints = New-Object System.Windows.Forms.Button
    $btnPickSalesPoints.Location = New-Object System.Drawing.Point(12, 98)
    $btnPickSalesPoints.Size = New-Object System.Drawing.Size(358, 28)
    $btnPickSalesPoints.Text = "SalesPoints selecteren..."
    $grpNew.Controls.Add($btnPickSalesPoints)

    $txtPicked = New-Object System.Windows.Forms.TextBox
    $txtPicked.Location = New-Object System.Drawing.Point(12, 134)
    $txtPicked.Size = New-Object System.Drawing.Size(358, 220)
    $txtPicked.Multiline = $true
    $txtPicked.ScrollBars = "Vertical"
    $txtPicked.ReadOnly = $true
    $grpNew.Controls.Add($txtPicked)

    function Render-Picked {
        $ids = @($dlg.Tag)
        $txtPicked.Clear()
        $txtPicked.AppendText("Geselecteerd: $($ids.Count)`r`n`r`n") | Out-Null
        foreach ($id in $ids) { $txtPicked.AppendText("- $id`r`n") | Out-Null }
    }

    $btnCreate = New-Object System.Windows.Forms.Button
    $btnCreate.Location = New-Object System.Drawing.Point(12, 366)
    $btnCreate.Size = New-Object System.Drawing.Size(358, 28)
    $btnCreate.Text = "Aanmaken"
    $grpNew.Controls.Add($btnCreate)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(734, 444)
    $btnClose.Size = New-Object System.Drawing.Size(96, 32)
    $btnClose.Text = "Sluiten"
    $btnClose.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($btnClose)

    # Events
    $btnPickSalesPoints.Add_Click({
        try {
            $salesPoints = Get-CachedSalesPoints
            if (-not $salesPoints -or $salesPoints.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("Geen salespoints beschikbaar.") | Out-Null
                return
            }

            $sel = Show-SalesPointMultiSelect -SalesPoints $salesPoints
            if (-not $sel -or $sel.Count -eq 0) { return }

            $ids = @(
                foreach ($x in $sel) {
                    if ($x -is [string]) { $x.Trim() }
                    elseif ($null -ne $x.PSObject.Properties['id']) { ([string]$x.id).Trim() }
                }
            ) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }

            $dlg.Tag = $ids
            Render-Picked
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fout bij selectie: $($_.Exception.Message)", "Fout") | Out-Null
        }
    })

    $btnCreate.Add_Click({
        try {
            $groupId = $txtId.Text.Trim()
            $name = $txtName.Text.Trim()
            $ids = @($dlg.Tag)

            if ([string]::IsNullOrWhiteSpace($groupId)) { throw "GroupId is verplicht." }
            if ([string]::IsNullOrWhiteSpace($name))    { throw "Naam is verplicht." }
            if (-not $ids -or $ids.Count -eq 0)         { throw "Selecteer minimaal 1 SalesPoint." }

            $btnCreate.Enabled = $false
            $dlg.UseWaitCursor = $true

            $newId = New-SalesPointGroupCore -Session $Session -GroupId $groupId -Name $name -SalesPointIds $ids

Write-Host "DEBUG Selector A: CreatedSalesPointGroups.Count = $($Session.CreatedSalesPointGroups.Count)"
            Refresh-List
Write-Host "DEBUG Selector B: CreatedSalesPointGroups.Count = $($Session.CreatedSalesPointGroups.Count)"

            $selectedId = $newId
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dlg.Close()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fout: $($_.Exception.Message)", "Fout") | Out-Null
        }
        finally {
            $dlg.UseWaitCursor = $false
            $btnCreate.Enabled = $true
        }
    })

    $btnUse.Add_Click({
        try {
            $spg = $lst.SelectedItem
            if ($null -eq $spg -or $null -eq $spg.PSObject.Properties['id']) {
                throw "Selecteer eerst een SalesPointGroup."
            }
            $selectedId = ([string]$spg.id).Trim()
            $dlg.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $dlg.Close()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fout: $($_.Exception.Message)", "Fout") | Out-Null
        }
    })

    Refresh-List
    Render-Picked

    $null = if ($OwnerForm) { $dlg.ShowDialog($OwnerForm) } else { $dlg.ShowDialog() }
    return $selectedId
}


function Show-NewAvailabilityGUI {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Session,
        [System.Windows.Forms.Form]$OwnerForm
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Availability aanmaken"
    $dlg.StartPosition = if ($OwnerForm) { "CenterParent" } else { "CenterScreen" }
    $dlg.Size = New-Object System.Drawing.Size(720, 460)
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false

	
	$chkKeepOpen = New-Object System.Windows.Forms.CheckBox
	$chkKeepOpen.Location = New-Object System.Drawing.Point(420, 400)
	$chkKeepOpen.Size = New-Object System.Drawing.Size(260, 20)
	$chkKeepOpen.Text = "Open houden na Opslaan"
	$chkKeepOpen.Checked = $true
	$dlg.Controls.Add($chkKeepOpen)
	


    # --- controls (id/name) ---
    $lblId = New-Object System.Windows.Forms.Label
    $lblId.Location = New-Object System.Drawing.Point(16, 16)
    $lblId.Size = New-Object System.Drawing.Size(120, 20)
    $lblId.Text = "ExternalId:"
    $dlg.Controls.Add($lblId)

    $txtId = New-Object System.Windows.Forms.TextBox
    $txtId.Location = New-Object System.Drawing.Point(150, 14)
    $txtId.Size = New-Object System.Drawing.Size(540, 20)
    $dlg.Controls.Add($txtId)

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Location = New-Object System.Drawing.Point(16, 48)
    $lblName.Size = New-Object System.Drawing.Size(120, 20)
    $lblName.Text = "Naam:"
    $dlg.Controls.Add($lblName)

    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(150, 46)
    $txtName.Size = New-Object System.Drawing.Size(540, 20)
    $dlg.Controls.Add($txtName)

    # --- SPG group ---
    $grpSpg = New-Object System.Windows.Forms.GroupBox
    $grpSpg.Location = New-Object System.Drawing.Point(16, 84)
    $grpSpg.Size = New-Object System.Drawing.Size(330, 280)
    $grpSpg.Text = "SalesPointGroup (uit sessie)"
    $dlg.Controls.Add($grpSpg)

    $lstSpg = New-Object System.Windows.Forms.ListBox
    $lstSpg.Location = New-Object System.Drawing.Point(12, 24)
    $lstSpg.Size = New-Object System.Drawing.Size(306, 218)
    $grpSpg.Controls.Add($lstSpg)

    $btnNewSpg = New-Object System.Windows.Forms.Button
    $btnNewSpg.Location = New-Object System.Drawing.Point(12, 248)
    $btnNewSpg.Size = New-Object System.Drawing.Size(306, 28)
    $btnNewSpg.Text = "SalesPointGroup selecteren/aanmaken..."
    $grpSpg.Controls.Add($btnNewSpg)

    # --- TP group ---
    $grpTp = New-Object System.Windows.Forms.GroupBox
    $grpTp.Location = New-Object System.Drawing.Point(360, 84)
    $grpTp.Size = New-Object System.Drawing.Size(330, 280)
    $grpTp.Text = "TimePeriod (uit sessie)"
    $dlg.Controls.Add($grpTp)

    $lstTimePeriods = New-Object System.Windows.Forms.ListBox
    $lstTimePeriods.Location = New-Object System.Drawing.Point(12, 24)
    $lstTimePeriods.Size = New-Object System.Drawing.Size(306, 218)
    $grpTp.Controls.Add($lstTimePeriods)

    $btnNewTp = New-Object System.Windows.Forms.Button
    $btnNewTp.Location = New-Object System.Drawing.Point(12, 248)
    $btnNewTp.Size = New-Object System.Drawing.Size(306, 28)
    $btnNewTp.Text = "Nieuwe TimePeriod aanmaken..."
    $grpTp.Controls.Add($btnNewTp)

    # --- bottom buttons ---
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Location = New-Object System.Drawing.Point(494, 380)
    $btnOk.Size = New-Object System.Drawing.Size(96, 32)
    $btnOk.Text = "Opslaan"
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(594, 380)
    $btnCancel.Size = New-Object System.Drawing.Size(96, 32)
    $btnCancel.Text = "Sluiten"
    $btnCancel.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($btnCancel)

    function Refresh-TimePeriods {
        $lstTimePeriods.Items.Clear()
        $tp = @()
        if ($Session.CreatedTimePeriods) { $tp = @($Session.CreatedTimePeriods) }
        foreach ($item in $tp) {
            $id = if ($item -is [string]) { $item } elseif ($item.PSObject.Properties['id']) { [string]$item.id } else { $null }
            if (-not [string]::IsNullOrWhiteSpace($id)) { [void]$lstTimePeriods.Items.Add($id.Trim()) }
        }
    }

	function Refresh-SalesPointGroups {
		$lstSpg.Items.Clear()
	
		if (-not $Session.PSObject.Properties.Match('CreatedSalesPointGroups')) {
			return
		}
	
		$spg = @()
		if ($Session.CreatedSalesPointGroups) { $spg = @($Session.CreatedSalesPointGroups) }
	
		foreach ($item in $spg) {
			$id =
				if ($item -is [string]) { $item }
				elseif ($null -ne $item.PSObject.Properties['id']) { [string]$item.id }
				else { $null }
	
			if (-not [string]::IsNullOrWhiteSpace($id)) {
				[void]$lstSpg.Items.Add($id.Trim())
			}
		}
	
		Write-Host "DEBUG Availability Refresh: lstSpg items = $($lstSpg.Items.Count)"
	}
    # Events
    $btnNewTp.Add_Click({
        try {
            Show-NewTimePeriodGUI -Session $Session -OwnerForm $dlg
            Refresh-TimePeriods
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fout bij TimePeriod: $($_.Exception.Message)") | Out-Null
        }
    })


	$btnNewSpg.Add_Click({
		try {
			$newId = Show-SalesPointGroupSelectorGUI -Session $Session -OwnerForm $dlg
	
			# ALTIJD refreshen (want sessie kan veranderd zijn)
			Refresh-SalesPointGroups
	
			# Als er een id is teruggegeven: selecteer hem
			if (-not [string]::IsNullOrWhiteSpace($newId)) {
				$idx = $lstSpg.Items.IndexOf([string]$newId)
				if ($idx -ge 0) {
					$lstSpg.SelectedIndex = $idx
					$lstSpg.TopIndex = $idx
				}
			}
		}
		catch {
			[System.Windows.Forms.MessageBox]::Show("Fout: $($_.Exception.Message)") | Out-Null
		}
	})
	
    $btnOk.Add_Click({
        try {
            $externalId = $txtId.Text.Trim()
            $name = $txtName.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($externalId)) { throw "ExternalId is verplicht." }
            if ([string]::IsNullOrWhiteSpace($name))       { throw "Naam is verplicht." }

            $spgId = if ($lstSpg.SelectedItem) { ([string]$lstSpg.SelectedItem).Trim() } else { $null }
            $tpId  = if ($lstTimePeriods.SelectedItem) { ([string]$lstTimePeriods.SelectedItem).Trim() } else { $null }

            if ([string]::IsNullOrWhiteSpace($spgId)) { throw "Selecteer een SalesPointGroup (of maak er één aan)." }
            if ([string]::IsNullOrWhiteSpace($tpId))  { throw "Selecteer een TimePeriod (of maak er één aan)." }

            $btnOk.Enabled = $false
            $dlg.UseWaitCursor = $true

            $id = New-AvailabilityCore -Session $Session -ExternalId $externalId -Name $name -SalesPointGroupId $spgId -TimePeriodId $tpId
            [System.Windows.Forms.MessageBox]::Show("Availability opgeslagen: $id", "Succes") | Out-Null
            $dlg.Close()
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fout: $($_.Exception.Message)", "Fout") | Out-Null
        }
        finally {
            $dlg.UseWaitCursor = $false
            $btnOk.Enabled = $true
        }
    })

    Refresh-SalesPointGroups
    Refresh-TimePeriods

    $null = if ($OwnerForm) { $dlg.ShowDialog($OwnerForm) } else { $dlg.ShowDialog() }
}








# ============================================================================
# CACHE MANAGEMENT
# ============================================================================

function Get-CachedStores {
    if ($null -eq $script:cachedStores) {
        Write-Host "Stores worden opgehaald..." -ForegroundColor Yellow
        $script:cachedStores = Get-OnboardingData -Endpoint "stores"
    }
    return $script:cachedStores
}

function Get-CachedSalesPoints {
    if ($null -eq $script:cachedSalesPoints) {
        Write-Host "SalesPoints worden opgehaald via stores..." -ForegroundColor Yellow
        $script:cachedSalesPoints = Get-SalesPointsFromStores
    }
    return $script:cachedSalesPoints
}

function Get-CachedTurnoverGroups {
    if ($null -eq $script:cachedTurnoverGroups) {
        Write-Host "TurnoverGroups worden opgehaald..." -ForegroundColor Yellow
        $script:cachedTurnoverGroups = Get-OnboardingData -Endpoint "turnovergroups"
    }
    return $script:cachedTurnoverGroups
}

function Get-CachedVatTariffs {
    if ($null -eq $script:cachedVatTariffs) {
        Write-Host "VatTariffs worden opgehaald..." -ForegroundColor Yellow
        $script:cachedVatTariffs = Get-OnboardingData -Endpoint "vattariffs"
    }
    return $script:cachedVatTariffs
}

function Get-CachedCustomers {
    if ($null -eq $script:cachedCustomers) {
        Write-Host "Customers worden opgehaald..." -ForegroundColor Yellow
        $script:cachedCustomers = Get-OnboardingData -Endpoint "customers"
    }
    return $script:cachedCustomers
}


function Get-CachedProducts {
    if ($null -eq $script:cachedProducts) {
        Write-Host "Products worden opgehaald..." -ForegroundColor Yellow
        $script:cachedProducts = Get-PimProducts

        # --- Alleen bruikbare producten (met id) behouden ---
        $incomingCount = if ($script:cachedProducts) { @($script:cachedProducts).Count } else { 0 }

        # Filter producten zonder id volledig weg
        $script:cachedProducts = @($script:cachedProducts | Where-Object {
            $_ -and
            -not [string]::IsNullOrWhiteSpace($_.name) -and
            -not [string]::IsNullOrWhiteSpace($_.id)
        })

        $filteredCount = if ($script:cachedProducts) { @($script:cachedProducts).Count } else { 0 }
        $removedCount = $incomingCount - $filteredCount

        if ($removedCount -gt 0) {
            Write-Log "$removedCount PIM product(en) genegeerd omdat 'id' ontbreekt (onbruikbaar voor selectie/koppeling)." -Level WARNING
        }

        # (optioneel) Deduplicate op id om dubbels te voorkomen
        $script:cachedProducts = @($script:cachedProducts | Group-Object id | ForEach-Object { $_.Group[0] })
    }

    return $script:cachedProducts
}



# Script gaat verder in deel 2...
# BooqCM-StamdataOnderhoud-Part2.ps1
# Dit bestand bevat de UI en selectie functies

# ============================================================================
# UI HELPER FUNCTIES
# ============================================================================

function Show-SelectionMenu {
    param(
        [string]$Title,
        [array]$Items,
        [string]$DisplayProperty,
        [string]$IdProperty = "id",
        [bool]$AllowMultiple = $false,
        [bool]$AllowCancel = $true
    )
    
    Clear-Host
    Write-Host "`n=== $Title ===" -ForegroundColor Cyan
    Write-Host ""
    
    if ($Items.Count -eq 0) {
        Write-Host "Geen items beschikbaar." -ForegroundColor Yellow
        Read-Host "`nDruk op Enter om door te gaan"
        return $null
    }
    
    for ($i = 0; $i -lt $Items.Count; $i++) {
        $item = $Items[$i]
        $displayValue = if ($item.PSObject.Properties[$DisplayProperty]) {
            $item.$DisplayProperty
        } else {
            $item.ToString()
        }
        
        $idValue = if ($item.PSObject.Properties[$IdProperty]) {
            " (ID: $($item.$IdProperty))"
        } else {
            ""
        }
        
        Write-Host "$($i + 1). $displayValue$idValue" -ForegroundColor White
    }
    
    Write-Host ""
    if ($AllowCancel) {
        Write-Host "C. Annuleren`n" -ForegroundColor Red
    }
    
    if ($AllowMultiple) {
        $prompt = "Selecteer items (komma-gescheiden)"
    } else {
        $prompt = "Selecteer een item"
    }
    
    $choice = Read-Host $prompt
    
    if ($AllowCancel -and ($choice -eq "C" -or $choice -eq "c")) {
        return $null
    }
    
    try {
        if ($AllowMultiple) {
            $indices = $choice -split ',' | ForEach-Object { [int]$_.Trim() - 1 }
            $selectedItems = @()
            foreach ($index in $indices) {
                if ($index -ge 0 -and $index -lt $Items.Count) {
                    $selectedItems += $Items[$index]
                }
            }
            return $selectedItems
        }
        else {
            $index = [int]$choice - 1
            if ($index -ge 0 -and $index -lt $Items.Count) {
                return $Items[$index]
            }
            else {
                Write-Host "Ongeldige selectie." -ForegroundColor Red
                Start-Sleep -Seconds 2
                return Show-SelectionMenu -Title $Title -Items $Items -DisplayProperty $DisplayProperty -IdProperty $IdProperty -AllowMultiple $AllowMultiple -AllowCancel $AllowCancel
            }
        }
    }
    catch {
        Write-Host "Ongeldige invoer." -ForegroundColor Red
        Start-Sleep -Seconds 2
        return Show-SelectionMenu -Title $Title -Items $Items -DisplayProperty $DisplayProperty -IdProperty $IdProperty -AllowMultiple $AllowMultiple -AllowCancel $AllowCancel
    }
}

function Read-InputWithValidation {
    param(
        [string]$Prompt,
        [string]$Type = "String",  # String, Int, Decimal, DateTime, Bool, Enum
        [bool]$Mandatory = $true,
        [array]$ValidValues = @(),
        [string]$DefaultValue = ""
    )
    
    $fullPrompt = $Prompt
    if (-not $Mandatory) {
        $fullPrompt += " (optioneel)"
    }
    if ($DefaultValue) {
        $fullPrompt += " [standaard: $DefaultValue]"
    }
    $fullPrompt += ": "
    
    while ($true) {
        $input = Read-Host $fullPrompt
        
        # Gebruik standaardwaarde als geen input en niet verplicht
        if ([string]::IsNullOrWhiteSpace($input)) {
            if ($DefaultValue) {
                return $DefaultValue
            }
            if (-not $Mandatory) {
                return $null
            }
            Write-Host "Dit veld is verplicht." -ForegroundColor Red
            continue
        }
        
        # Validatie op basis van type
        try {
            switch ($Type) {
                "String" {
                    if ($ValidValues.Count -gt 0 -and $input -notin $ValidValues) {
                        Write-Host "Ongeldige waarde. Toegestane waardes: $($ValidValues -join ', ')" -ForegroundColor Red
                        continue
                    }
                    return $input
                }
                "Int" {
                    return [int]$input
                }
                "Decimal" {
                    return [decimal]$input
                }
                "DateTime" {
                    return [datetime]$input
                }
                "Bool" {
                    if ($input -match '^(true|yes|ja|1)$') { return $true }
                    if ($input -match '^(false|no|nee|0)$') { return $false }
                    Write-Host "Ongeldige boolean. Gebruik: true/false, yes/no, ja/nee, 1/0" -ForegroundColor Red
                    continue
                }
                "Enum" {
                    if ($ValidValues.Count -gt 0 -and $input -notin $ValidValues) {
                        Write-Host "Ongeldige waarde. Toegestane waardes: $($ValidValues -join ', ')" -ForegroundColor Red
                        continue
                    }
                    return $input
                }
                default {
                    return $input
                }
            }
        }
        catch {
            Write-Host "Ongeldige invoer voor type $Type." -ForegroundColor Red
            continue
        }
    }
}

# ============================================================================
# HISTORY MANAGEMENT
# ============================================================================

function Add-ToHistory {
    param(
        [PSCustomObject]$Session,
        [string]$Action,
        [string]$EntityType,
        [string]$EntityId,
        [hashtable]$Parameters,
        [object]$Result
    )
    
    $historyEntry = [PSCustomObject]@{
        Timestamp = Get-Date
        Action = $Action
        EntityType = $EntityType
        EntityId = $EntityId
        Parameters = $Parameters
        Result = $Result
        Success = $true
    }
    
    $Session.History += $historyEntry
    
    # Voeg toe aan relevante lijst
    switch ($EntityType) {
        "TimePeriod" {
            if ($Action -eq "Create") {
                $Session.CreatedTimePeriods += $EntityId
            }
        }
        "Availability" {
            if ($Action -eq "Create") {
                if ($EntityId -notin $Session.CreatedAvailabilities) {
                    $Session.CreatedAvailabilities += $EntityId
                    if ($DebugMode) {
                        Write-Log "Availability $EntityId toegevoegd aan sessie (via Add-ToHistory)" -level DEBUG
                    }
                } else {
                    if ($DebugMode) {
                        Write-Log "Availability $EntityId was al in sessie" -level DEBUG
                    }
                }
            }
        }
        "Customer" {
            if ($Action -eq "Create") {
                $Session.CreatedCustomers += $EntityId
            }
        }
        "Promotion" {
            if ($Action -eq "Create") {
                $Session.CreatedPromotions += $EntityId
            }
        }
    }
    
    # FIX v1.3: Suppress return waarde van Save-Session om array problemen te voorkomen
    [void](Save-Session -Session $Session)
}

function Show-History {
    param([PSCustomObject]$Session)
    
    Clear-Host
    Write-Host "`n=== SESSIE HISTORY ===" -ForegroundColor Cyan
    Write-Host "Sessie: $($Session.Name)" -ForegroundColor Yellow
    Write-Host "Totaal aantal acties: $($Session.History.Count)`n" -ForegroundColor Yellow
    
    if ($Session.History.Count -eq 0) {
        Write-Host "Geen acties uitgevoerd in deze sessie." -ForegroundColor Gray
        Read-Host "`nDruk op Enter om door te gaan"
        return
    }
    
    for ($i = 0; $i -lt $Session.History.Count; $i++) {
        $entry = $Session.History[$i]
        Write-Host "[$($i + 1)] $($entry.Timestamp.ToString('yyyy-MM-dd HH:mm:ss'))" -ForegroundColor Cyan
        Write-Host "    Actie: $($entry.Action) $($entry.EntityType)" -ForegroundColor White
        Write-Host "    Entity ID: $($entry.EntityId)" -ForegroundColor Gray
        
        if ($entry.Parameters -and $entry.Parameters.Count -gt 0) {
            Write-Host "    Parameters:" -ForegroundColor Gray
            foreach ($key in $entry.Parameters.Keys) {
                $value = $entry.Parameters[$key]
                if ($value -is [array]) {
                    Write-Host "      $key`: [array met $($value.Count) items]" -ForegroundColor DarkGray
                } elseif ($value -is [hashtable] -or $value -is [PSCustomObject]) {
                    Write-Host "      $key`: [object]" -ForegroundColor DarkGray
                } else {
                    Write-Host "      $key`: $value" -ForegroundColor DarkGray
                }
            }
        }
        Write-Host ""
    }
    
    Write-Host "`n1. Actie herhalen/aanpassen" -ForegroundColor Yellow
    Write-Host "2. Terug naar hoofdmenu`n" -ForegroundColor White
    
    $choice = Read-Host "Maak uw keuze"
    
    switch ($choice) {
        "1" {
            $actionNr = Read-Host "Welke actie wilt u herhalen? (nummer)"
            try {
                $index = [int]$actionNr - 1
                if ($index -ge 0 -and $index -lt $Session.History.Count) {
                    Repeat-HistoryAction -Session $Session -HistoryEntry $Session.History[$index]
                } else {
                    Write-Host "Ongeldige actie nummer." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    Show-History -Session $Session
                }
            }
            catch {
                Write-Host "Ongeldige invoer." -ForegroundColor Red
                Start-Sleep -Seconds 2
                Show-History -Session $Session
            }
        }
        "2" {
            return
        }
        default {
            Show-History -Session $Session
        }
    }
}

function Repeat-HistoryAction {
    param(
        [PSCustomObject]$Session,
        [PSCustomObject]$HistoryEntry
    )
    
    Clear-Host
    Write-Host "`n=== ACTIE HERHALEN/AANPASSEN ===" -ForegroundColor Cyan
    Write-Host "`nOriginele actie:" -ForegroundColor Yellow
    Write-Host "  $($HistoryEntry.Action) $($HistoryEntry.EntityType)" -ForegroundColor White
    Write-Host "  Entity ID: $($HistoryEntry.EntityId)" -ForegroundColor Gray
    
    Write-Host "`n1. Herhalen met zelfde waardes" -ForegroundColor Green
    Write-Host "2. Herhalen met aangepaste waardes" -ForegroundColor Yellow
    Write-Host "3. Annuleren`n" -ForegroundColor Red
    
    $choice = Read-Host "Maak uw keuze"
    
    switch ($choice) {
        "1" {
            # Herhaal met originele parameters
            Write-Host "`nActie wordt herhaald..." -ForegroundColor Yellow
            
            $entityType = $HistoryEntry.EntityType
            $action = $HistoryEntry.Action
            $params = $HistoryEntry.Parameters
            
            if ($action -eq "Create" -and $entityType -eq "Promotion") {
                # Maak nieuwe ID aan
                $newId = [guid]::NewGuid().ToString()
                Write-Host "Nieuwe ID gegenereerd: $newId" -ForegroundColor Cyan
                
                try {
                    Invoke-CMApi -Endpoint "promotions/$newId" -Method "PUT" -Body $params.Body
                    Write-Host "Promotie succesvol aangemaakt!" -ForegroundColor Green
                    Add-ToHistory -Session $Session -Action "Create" -EntityType "Promotion" -EntityId $newId -Parameters $params -Result $null
                }
                catch {
                    Write-Host "Fout bij aanmaken promotie: $_" -ForegroundColor Red
                }
            }
            
            Read-Host "`nDruk op Enter om door te gaan"
        }
        "2" {
            Write-Host "`nDeze functionaliteit wordt nog geïmplementeerd..." -ForegroundColor Yellow
            Read-Host "`nDruk op Enter om door te gaan"
        }
        "3" {
            return
        }
    }
}

# Vervolg in Part 3...
# BooqCM-StamdataOnderhoud-Part3.ps1
# Dit bestand bevat de core business logica

# ============================================================================
# TIMEPERIOD FUNCTIES
# ============================================================================


	
function To-UtcLiteralIso {
    param([Parameter(Mandatory)][datetime]$Dt)
    # Neem de ingevoerde kloktijd en label die als UTC (geen conversie!)
    $asUtc = [DateTime]::SpecifyKind($Dt, [DateTimeKind]::Utc)
    return $asUtc.ToString("o")  # eindigt met Z
}


function New-TimePeriodCore {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Session,

        [Parameter(Mandatory)]
        [string]$ExternalId,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [datetime]$StartsAt,

        [Parameter(Mandatory)]
        [datetime]$EndsAt,

        [ValidateSet("NONE","DAY","WEEK","MONTH","YEAR")]
        [string]$RepetitionType = "NONE",

        [datetime]$RepetitionEndsAt,

        [int]$RepetitionInterval
    )


Write-Log "tmp is dit DebugMode?   ; ($DebugMode) " 

if ($DebugMode) {
    Write-Log ">>> New-TimePeriodCore (debug) - entered" -level DEBUG
    Write-Log ("Bound params: {0}" -f (($PSBoundParameters.Keys | Sort-Object) -join ", ")) -level DEBUG
    if ($PSBoundParameters.ContainsKey('RepetitionType')) {
        Write-Log ("RepetitionType={0}" -f $RepetitionType) -level DEBUG
    }
    if ($PSBoundParameters.ContainsKey('RepetitionInterval')) {
        Write-Log ("RepetitionInterval={0}" -f $RepetitionInterval) -level DEBUG
    } else {
        Write-Log "RepetitionInterval not bound" -level DEBUG
    }
    if ($PSBoundParameters.ContainsKey('RepetitionEndsAt')) {
        Write-Log ("RepetitionEndsAt={0:o}" -f $RepetitionEndsAt) -level DEBUG
    } else {
        Write-Log "RepetitionEndsAt not bound" -level DEBUG
    }
}



    $ExternalId = $ExternalId.Trim()
    $Name = $Name.Trim()

    if ([string]::IsNullOrWhiteSpace($ExternalId)) { throw "ExternalId is verplicht." }
    if ([string]::IsNullOrWhiteSpace($Name)) { throw "Naam is verplicht." }
    if ($EndsAt -le $StartsAt) { throw "Einde (endsAt) moet na start (startsAt) liggen." }

    $rep = if ([string]::IsNullOrWhiteSpace($RepetitionType)) { "NONE" } else { $RepetitionType }

    if ($rep -ne "NONE") {
        if (-not $PSBoundParameters.ContainsKey('RepetitionInterval')) {
            throw "repetitionInterval is vereist wanneer repetitionType <> NONE."
        }
        if ($RepetitionInterval -lt 1) {
            throw "repetitionInterval moet minimaal 1 zijn."
        }

        if (-not $PSBoundParameters.ContainsKey('RepetitionEndsAt')) {
            throw "repetitionEndsAt is vereist wanneer repetitionType <> NONE."
        }
        if ($RepetitionEndsAt -le $EndsAt) {
            throw "repetitionEndsAt moet na endsAt liggen."
        }
    }

    $body = @{
        name     = $Name
        ## startsAt = $StartsAt.ToUniversalTime().ToString("o")
        ## endsAt   = $EndsAt.ToUniversalTime().ToString("o")
		startsAt = (To-UtcLiteralIso $StartsAt)		
		endsAt   = (To-UtcLiteralIso $EndsAt)
        repetitionType = $rep
    }

    if ($rep -ne "NONE") {
        $body['repetitionInterval'] = [int]$RepetitionInterval
        # $body['repetitionEndsAt']   = $RepetitionEndsAt.ToUniversalTime().ToString("o")
        $body['repetitionEndsAt']   = (To-UtcLiteralIso $RepetitionEndsAt)	
    }

	if ($DebugMode) {
		$dbgJson = ($body | ConvertTo-Json -Depth 20 -Compress)
		Write-Log ("New-TimePeriodCore body (debug): {0}" -f $dbgJson) -level DEBUG
	}
	
    $result = Invoke-CMApi -Endpoint "timeperiods/$ExternalId" -Method "PUT" -Body $body
    Add-ToHistory -Session $Session -Action "Create" -EntityType "TimePeriod" -EntityId $ExternalId -Parameters @{ Body = $body } -Result $result

    if (-not $Session.CreatedTimePeriods) { $Session.CreatedTimePeriods = @() }
    if ($Session.CreatedTimePeriods -notcontains $ExternalId) { $Session.CreatedTimePeriods += $ExternalId }

    [void](Save-Session -Session $Session)

    return $ExternalId
}

		
function Show-NewTimePeriodGUI {
		param(
			[Parameter(Mandatory)]
			[PSCustomObject]$Session,
			[System.Windows.Forms.Form]$OwnerForm
		)
	
		Add-Type -AssemblyName System.Windows.Forms
		Add-Type -AssemblyName System.Drawing
	
		$dlg = New-Object System.Windows.Forms.Form
		$dlg.Text = "TimePeriod aanmaken"
		$dlg.StartPosition = "CenterParent"
		$dlg.Size = New-Object System.Drawing.Size(620, 460)
		$dlg.FormBorderStyle = "FixedDialog"
		$dlg.MaximizeBox = $false
		$dlg.MinimizeBox = $false
		$dlg.ShowInTaskbar = $false
	
		# ExternalId
		$lblId = New-Object System.Windows.Forms.Label
		$lblId.Location = New-Object System.Drawing.Point(16, 16)
		$lblId.Size = New-Object System.Drawing.Size(120, 20)
		$lblId.Text = "ExternalId:"
		$dlg.Controls.Add($lblId)
	
		$txtId = New-Object System.Windows.Forms.TextBox
		$txtId.Location = New-Object System.Drawing.Point(150, 14)
		$txtId.Size = New-Object System.Drawing.Size(440, 20)
		$dlg.Controls.Add($txtId)
	
		# Name
		$lblName = New-Object System.Windows.Forms.Label
		$lblName.Location = New-Object System.Drawing.Point(16, 50)
		$lblName.Size = New-Object System.Drawing.Size(120, 20)
		$lblName.Text = "Naam:"
		$dlg.Controls.Add($lblName)
	
		$txtName = New-Object System.Windows.Forms.TextBox
		$txtName.Location = New-Object System.Drawing.Point(150, 48)
		$txtName.Size = New-Object System.Drawing.Size(440, 20)
		$dlg.Controls.Add($txtName)
	
		# RepetitionType
		$lblRep = New-Object System.Windows.Forms.Label
		$lblRep.Location = New-Object System.Drawing.Point(16, 84)
		$lblRep.Size = New-Object System.Drawing.Size(120, 20)
		$lblRep.Text = "RepetitionType:"
		$dlg.Controls.Add($lblRep)
	
		$cmbRep = New-Object System.Windows.Forms.ComboBox
		$cmbRep.Location = New-Object System.Drawing.Point(150, 82)
		$cmbRep.Size = New-Object System.Drawing.Size(200, 20)
		$cmbRep.DropDownStyle = "DropDownList"
		@("NONE","DAY","WEEK","MONTH","YEAR") | ForEach-Object { [void]$cmbRep.Items.Add($_) }
		$cmbRep.SelectedItem = "NONE"
		$dlg.Controls.Add($cmbRep)
	
		# Start date + time
		$lblStart = New-Object System.Windows.Forms.Label
		$lblStart.Location = New-Object System.Drawing.Point(16, 124)
		$lblStart.Size = New-Object System.Drawing.Size(120, 20)
		$lblStart.Text = "Start:"
		$dlg.Controls.Add($lblStart)
	
		$dtStartDate = New-Object System.Windows.Forms.DateTimePicker
		$dtStartDate.Location = New-Object System.Drawing.Point(150, 120)
		$dtStartDate.Size = New-Object System.Drawing.Size(200, 20)
		$dtStartDate.Format = "Short"
		$dlg.Controls.Add($dtStartDate)
	
		$dtStartTime = New-Object System.Windows.Forms.DateTimePicker
		$dtStartTime.Location = New-Object System.Drawing.Point(360, 120)
		$dtStartTime.Size = New-Object System.Drawing.Size(120, 20)
		$dtStartTime.Format = "Time"
		$dtStartTime.ShowUpDown = $true
		$dlg.Controls.Add($dtStartTime)
	
		# End date + time
		$lblEnd = New-Object System.Windows.Forms.Label
		$lblEnd.Location = New-Object System.Drawing.Point(16, 162)
		$lblEnd.Size = New-Object System.Drawing.Size(120, 20)
		$lblEnd.Text = "Einde:"
		$dlg.Controls.Add($lblEnd)
	
		$dtEndDate = New-Object System.Windows.Forms.DateTimePicker
		$dtEndDate.Location = New-Object System.Drawing.Point(150, 158)
		$dtEndDate.Size = New-Object System.Drawing.Size(200, 20)
		$dtEndDate.Format = "Short"
		$dlg.Controls.Add($dtEndDate)
	
		$dtEndTime = New-Object System.Windows.Forms.DateTimePicker
		$dtEndTime.Location = New-Object System.Drawing.Point(360, 158)
		$dtEndTime.Size = New-Object System.Drawing.Size(120, 20)
		$dtEndTime.Format = "Time"
		$dtEndTime.ShowUpDown = $true
		$dlg.Controls.Add($dtEndTime)
	
		# repetitionInterval
		$lblInterval = New-Object System.Windows.Forms.Label
		$lblInterval.Location = New-Object System.Drawing.Point(16, 200)
		$lblInterval.Size = New-Object System.Drawing.Size(120, 20)
		$lblInterval.Text = "Interval:"
		$dlg.Controls.Add($lblInterval)
	
		$numInterval = New-Object System.Windows.Forms.NumericUpDown
		$numInterval.Location = New-Object System.Drawing.Point(150, 198)
		$numInterval.Size = New-Object System.Drawing.Size(120, 20)
		$numInterval.Minimum = 1
		$numInterval.Maximum = 365
		$numInterval.Value = 1
		$dlg.Controls.Add($numInterval)
	
		# repetitionEndsAt (date+time)
		$lblRepEnd = New-Object System.Windows.Forms.Label
		$lblRepEnd.Location = New-Object System.Drawing.Point(16, 238)
		$lblRepEnd.Size = New-Object System.Drawing.Size(120, 20)
		$lblRepEnd.Text = "Repeat tot:"
		$dlg.Controls.Add($lblRepEnd)
	
		$dtRepEndDate = New-Object System.Windows.Forms.DateTimePicker
		$dtRepEndDate.Location = New-Object System.Drawing.Point(150, 234)
		$dtRepEndDate.Size = New-Object System.Drawing.Size(200, 20)
		$dtRepEndDate.Format = "Short"
		$dlg.Controls.Add($dtRepEndDate)
	
		$dtRepEndTime = New-Object System.Windows.Forms.DateTimePicker
		$dtRepEndTime.Location = New-Object System.Drawing.Point(360, 234)
		$dtRepEndTime.Size = New-Object System.Drawing.Size(120, 20)
		$dtRepEndTime.Format = "Time"
		$dtRepEndTime.ShowUpDown = $true
		$dlg.Controls.Add($dtRepEndTime)
	
		# Defaults
		$dtStartTime.Value = [datetime]::Today.Date.AddHours(0).AddMinutes(0)
		$dtEndTime.Value   = [datetime]::Today.Date.AddHours(23).AddMinutes(59)
		$dtEndDate.Value   = $dtStartDate.Value.Date
		$dtRepEndDate.Value = $dtStartDate.Value.Date.AddMonths(1)
		$dtRepEndTime.Value = $dtEndTime.Value
	
		function Update-RepetitionControls {
			$rep = [string]$cmbRep.SelectedItem
			$isRepeating = ($rep -and $rep -ne "NONE")
	
			$lblInterval.Enabled = $isRepeating
			$numInterval.Enabled = $isRepeating
	
			$lblRepEnd.Enabled = $isRepeating
			$dtRepEndDate.Enabled = $isRepeating
			$dtRepEndTime.Enabled = $isRepeating
		}
	
		$cmbRep.Add_SelectedIndexChanged({ Update-RepetitionControls })
		Update-RepetitionControls
	
		# Buttons
		$btnOk = New-Object System.Windows.Forms.Button
		$btnOk.Location = New-Object System.Drawing.Point(394, 360)
		$btnOk.Size = New-Object System.Drawing.Size(96, 32)
		$btnOk.Text = "Opslaan"
		$dlg.Controls.Add($btnOk)
	
		$btnCancel = New-Object System.Windows.Forms.Button
		$btnCancel.Location = New-Object System.Drawing.Point(494, 360)
		$btnCancel.Size = New-Object System.Drawing.Size(96, 32)
		$btnCancel.Text = "Sluiten"
		$btnCancel.Add_Click({ $dlg.Close() })
		$dlg.Controls.Add($btnCancel)
	
		$btnOk.Add_Click({
			try {
				$externalId = $txtId.Text.Trim()
				if ([string]::IsNullOrWhiteSpace($externalId)) {
					[System.Windows.Forms.MessageBox]::Show("ExternalId is verplicht.") | Out-Null
					return
				}
	
				$name = $txtName.Text.Trim()
				if ([string]::IsNullOrWhiteSpace($name)) {
					[System.Windows.Forms.MessageBox]::Show("Naam is verplicht.") | Out-Null
					return
				}
	
				$startsAt = $dtStartDate.Value.Date.AddHours($dtStartTime.Value.Hour).AddMinutes($dtStartTime.Value.Minute)
				$endsAt   = $dtEndDate.Value.Date.AddHours($dtEndTime.Value.Hour).AddMinutes($dtEndTime.Value.Minute)
				if ($endsAt -le $startsAt) {
					[System.Windows.Forms.MessageBox]::Show("Einde moet na start liggen.") | Out-Null
					return
				}
	
				$rep = [string]$cmbRep.SelectedItem
				if ([string]::IsNullOrWhiteSpace($rep)) { $rep = "NONE" }
	
				$btnOk.Enabled = $false
				$dlg.UseWaitCursor = $true
	
				$coreParams = @{
					Session = $Session
					ExternalId = $externalId
					Name = $name
					StartsAt = $startsAt
					EndsAt = $endsAt
					RepetitionType = $rep
				}
	
				if ($rep -ne "NONE") {
					$repEndsAt = $dtRepEndDate.Value.Date.AddHours($dtRepEndTime.Value.Hour).AddMinutes($dtRepEndTime.Value.Minute)
					$coreParams.RepetitionEndsAt = $repEndsAt
					$coreParams.RepetitionInterval = [int]$numInterval.Value
				}
	
				$null = New-TimePeriodCore @coreParams
	
				[System.Windows.Forms.MessageBox]::Show("TimePeriod opgeslagen: $externalId", "Succes") | Out-Null
				$dlg.Close()
			}
			catch {
				[System.Windows.Forms.MessageBox]::Show("Fout: $($_.Exception.Message)", "Fout") | Out-Null
			}
			finally {
				$dlg.UseWaitCursor = $false
				$btnOk.Enabled = $true
			}
		})
	
		$null = $dlg.ShowDialog($OwnerForm)
	}
	

##	function New-TimePeriod {
##	    param([PSCustomObject]$Session)
##	    
##	    $retry = $true
##	    
##	    while ($retry) {
##	        Clear-Host
##	        Write-Host "`n=== NIEUWE TIMEPERIOD AANMAKEN ===" -ForegroundColor Cyan
##	        Write-Host ""
##	        
##	        try {
##	            $timePeriodId = Read-InputWithValidation -Prompt "Geef een unieke ID voor de TimePeriod" -Type "String" -Mandatory $true
##	            $name = Read-InputWithValidation -Prompt "Naam van de TimePeriod" -Type "String" -Mandatory $true
##	            
##	            Write-Host "`nStart datum/tijd (formaat: yyyy-MM-dd HH:mm)" -ForegroundColor Yellow
##	            $startsAtStr = Read-InputWithValidation -Prompt "Start datum/tijd" -Type "String" -Mandatory $true
##	            
##	            # Valideer datum formaat
##	            try {
##	                $startsAt = [datetime]::ParseExact($startsAtStr, "yyyy-MM-dd HH:mm", $null).ToString("yyyy-MM-ddTHH:mm:ss")
##	            }
##	            catch {
##	                Write-Host "`n[FOUT] Ongeldige datum/tijd. Gebruik formaat: yyyy-MM-dd HH:mm (bijv. 2025-10-27 14:30)" -ForegroundColor Red
##	                Read-Host "Druk op Enter om opnieuw te proberen"
##	                continue
##	            }
##	            
##	            Write-Host "`nEind datum/tijd (formaat: yyyy-MM-dd HH:mm)" -ForegroundColor Yellow
##	            $endsAtStr = Read-InputWithValidation -Prompt "Eind datum/tijd" -Type "String" -Mandatory $true
##	            
##	            # Valideer datum formaat
##	            try {
##	                $endsAt = [datetime]::ParseExact($endsAtStr, "yyyy-MM-dd HH:mm", $null).ToString("yyyy-MM-ddTHH:mm:ss")
##	            }
##	            catch {
##	                Write-Host "`n[FOUT] Ongeldige datum/tijd. Gebruik formaat: yyyy-MM-dd HH:mm (bijv. 2025-11-02 23:59)" -ForegroundColor Red
##	                Read-Host "Druk op Enter om opnieuw te proberen"
##	                continue
##	            }
##	            
##	            # Controleer of eind na start is
##	            $startDate = [datetime]::ParseExact($startsAtStr, "yyyy-MM-dd HH:mm", $null)
##	            $endDate = [datetime]::ParseExact($endsAtStr, "yyyy-MM-dd HH:mm", $null)
##	            
##	            if ($endDate -le $startDate) {
##	                Write-Host "`n[FOUT] Eind datum/tijd moet na start datum/tijd liggen!" -ForegroundColor Red
##	                Read-Host "Druk op Enter om opnieuw te proberen"
##	                continue
##	            }
##	            
##	            Write-Host "`nHerhaling type:" -ForegroundColor Yellow
##	            Write-Host "1. NONE (geen herhaling)"
##	            Write-Host "2. DAY (dagelijks)"
##	            Write-Host "3. WEEK (wekelijks)"
##	            Write-Host "4. MONTH (maandelijks)"
##	            Write-Host "5. YEAR (jaarlijks)"
##	            
##	            $repChoice = Read-Host "Keuze"
##	            $repetitionType = switch ($repChoice) {
##	                "1" { "NONE" }
##	                "2" { "DAY" }
##	                "3" { "WEEK" }
##	                "4" { "MONTH" }
##	                "5" { "YEAR" }
##	                default { "NONE" }
##	            }
##	            
##	            $body = @{
##	                name = $name
##	                startsAt = $startsAt
##	                endsAt = $endsAt
##	                repetitionType = $repetitionType
##	            }
##	            
##	            if ($repetitionType -ne "NONE") {
##	                $repetitionInterval = Read-InputWithValidation -Prompt "Herhaling interval (bijv. 1 voor elke dag/week/maand)" -Type "Int" -Mandatory $false -DefaultValue "1"
##	                if ($repetitionInterval) {
##	                    $body['repetitionInterval'] = [string]$repetitionInterval
##	                }
##	                
##	                Write-Host "`n[BELANGRIJK] Eind datum herhaling is VERPLICHT voor herhalende periods!" -ForegroundColor Yellow
##	                Write-Host "Formaat: yyyy-MM-dd HH:mm" -ForegroundColor Yellow
##	                $repetitionEndsAtStr = Read-InputWithValidation -Prompt "Eind datum herhaling" -Type "String" -Mandatory $true
##	                
##	                # Valideer datum formaat
##	                try {
##	                    $repetitionEndsAt = [datetime]::ParseExact($repetitionEndsAtStr, "yyyy-MM-dd HH:mm", $null).ToString("yyyy-MM-ddTHH:mm:ss")
##	                    $body['repetitionEndsAt'] = $repetitionEndsAt
##	                }
##	                catch {
##	                    Write-Host "`n[FOUT] Ongeldige datum/tijd. Gebruik formaat: yyyy-MM-dd HH:mm" -ForegroundColor Red
##	                    Read-Host "Druk op Enter om opnieuw te proberen"
##	                    continue
##	                }
##	                
##	                # Controleer of repetition ends na period ends is
##	                $repEndDate = [datetime]::ParseExact($repetitionEndsAtStr, "yyyy-MM-dd HH:mm", $null)
##	                if ($repEndDate -le $endDate) {
##	                    Write-Host "`n[WAARSCHUWING] Herhaling eindigt voor of op de period eind datum." -ForegroundColor Yellow
##	                    Write-Host "Dit betekent dat de herhaling mogelijk niet effectief is." -ForegroundColor Yellow
##	                    $continue = Read-InputWithValidation -Prompt "Toch doorgaan? (ja/nee)" -Type "String" -Mandatory $true
##	                    if ($continue -notmatch "^(ja|yes|j|y)$") {
##	                        Read-Host "Druk op Enter om opnieuw te proberen"
##	                        continue
##	                    }
##	                }
##	            }
##	            
##	            # Toon samenvatting
##	            Write-Host "`n=== SAMENVATTING TIMEPERIOD ===" -ForegroundColor Cyan
##	            Write-Host "ID: $timePeriodId" -ForegroundColor White
##	            Write-Host "Naam: $name" -ForegroundColor White
##	            Write-Host "Start: $startsAtStr" -ForegroundColor White
##	            Write-Host "Eind: $endsAtStr" -ForegroundColor White
##	            Write-Host "Herhaling: $repetitionType" -ForegroundColor White
##	            if ($repetitionType -ne "NONE") {
##	                Write-Host "Herhaling interval: $($body['repetitionInterval'])" -ForegroundColor White
##	                Write-Host "Herhaling eindigt: $repetitionEndsAtStr" -ForegroundColor White
##	            }
##	            Write-Host "===============================" -ForegroundColor Cyan
##	            
##	            $confirm = Read-InputWithValidation -Prompt "`nTimePeriod aanmaken? (ja/nee)" -Type "String" -Mandatory $true
##	            if ($confirm -notmatch "^(ja|yes|j|y)$") {
##	                Write-Host "TimePeriod niet aangemaakt." -ForegroundColor Yellow
##	                $retry = $false
##	                Read-Host "`nDruk op Enter om door te gaan"
##	                return $null
##	            }
##	            
##	            Write-Host "`nTimePeriod wordt aangemaakt..." -ForegroundColor Yellow
##	            
##	            $result = Invoke-CMApi -Endpoint "timeperiods/$timePeriodId" -Method "PUT" -Body $body
##	            Write-Host "`n[OK] TimePeriod succesvol aangemaakt!" -ForegroundColor Green
##	            
##	            Add-ToHistory -Session $Session -Action "Create" -EntityType "TimePeriod" -EntityId $timePeriodId -Parameters @{ Body = $body } -Result $result
##	            
##	            Read-Host "`nDruk op Enter om door te gaan"
##	            return $timePeriodId
##	        }
##	        catch {
##	            Write-Host "`n[FOUT] Fout bij aanmaken TimePeriod:" -ForegroundColor Red
##	            Write-Host "$_" -ForegroundColor Red
##	            
##	            # Parse error message voor specifieke feedback
##	            $errorMsg = $_.ToString()
##	            if ($errorMsg -match "repetitionEndsAt.*missing") {
##	                Write-Host "`nProbleem: De API vereist een eind datum voor de herhaling." -ForegroundColor Yellow
##	                Write-Host "Dit is verplicht wanneer je een herhaling type kiest (niet NONE)." -ForegroundColor Yellow
##	            }
##	            elseif ($errorMsg -match "ParseExact") {
##	                Write-Host "`nProbleem: Datum/tijd formaat is incorrect." -ForegroundColor Yellow
##	                Write-Host "Gebruik: yyyy-MM-dd HH:mm (bijvoorbeeld: 2025-10-27 14:30)" -ForegroundColor Yellow
##	            }
##	            
##	            Write-Host "`nWat wilt u doen?" -ForegroundColor Cyan
##	            Write-Host "1. Opnieuw proberen" -ForegroundColor Green
##	            Write-Host "2. Annuleren en terug naar menu" -ForegroundColor Red
##	            
##	            $choice = Read-Host "`nKeuze"
##	            
##	            if ($choice -eq "2") {
##	                $retry = $false
##	                return $null
##	            }
##	            # Anders loopt de while loop opnieuw
##	        }
##	    }
##	    
##	    return $null
##	}
##	
##	function Select-TimePeriod {
##	    param(
##	        [PSCustomObject]$Session,
##	        [string]$Purpose = "Selecteer een TimePeriod"
##	    )
##	    
##	    # Eerst kijken of er TimePeriods in de sessie zijn
##	    $sessionTimePeriods = @()
##	    if ($Session.CreatedTimePeriods.Count -gt 0) {
##	        foreach ($tpId in $Session.CreatedTimePeriods) {
##	            $sessionTimePeriods += [PSCustomObject]@{
##	                id = $tpId
##	                name = "TimePeriod uit sessie"
##	                source = "session"
##	            }
##	        }
##	    }
##	    
##	    Clear-Host
##	    Write-Host "`n=== $Purpose ===" -ForegroundColor Cyan
##	    Write-Host ""
##	    $tpColor = if ($sessionTimePeriods.Count -gt 0) { "Green" } else { "Gray" }
##	    $tpText = $("1. TimePeriod uit deze sessie selecteren (" + $sessionTimePeriods.Count + " beschikbaar)")
##	    Write-Host $tpText -ForegroundColor $tpColor
##	    Write-Host "2. Nieuwe TimePeriod aanmaken" -ForegroundColor Yellow
##	    Write-Host "3. Handmatig ID invoeren" -ForegroundColor White
##	    Write-Host "C. Annuleren`n" -ForegroundColor Red
##	    
##	    $choice = Read-Host "Maak uw keuze"
##	    
##	    switch ($choice) {
##	        "1" {
##	            if ($sessionTimePeriods.Count -eq 0) {
##	                Write-Host "Geen TimePeriods in deze sessie." -ForegroundColor Yellow
##	                Start-Sleep -Seconds 2
##	                return Select-TimePeriod -Session $Session -Purpose $Purpose
##	            }
##	            
##	            $selected = Show-SelectionMenu -Title "Selecteer TimePeriod uit sessie" -Items $sessionTimePeriods -DisplayProperty "id" -IdProperty "id" -AllowCancel $true
##	            if ($selected) {
##	                return $selected.id
##	            }
##	            return Select-TimePeriod -Session $Session -Purpose $Purpose
##	        }
##	        "2" {
##	            $newId = New-TimePeriod -Session $Session
##	            if ($newId) {
##	                return $newId
##	            }
##	            return Select-TimePeriod -Session $Session -Purpose $Purpose
##	        }
##	        "3" {
##	            $manualId = Read-InputWithValidation -Prompt "Geef TimePeriod ID" -Type "String" -Mandatory $true
##	            return $manualId
##	        }
##	        "C" {
##	            return $null
##	        }
##	        "c" {
##	            return $null
##	        }
##	        default {
##	            return Select-TimePeriod -Session $Session -Purpose $Purpose
##	        }
##	    }
##	}
##	

# ============================================================================
# AVAILABILITY FUNCTIES
# ============================================================================

function New-Availability {
    param([PSCustomObject]$Session)
    
    Clear-Host
    Write-Host "`n=== NIEUWE AVAILABILITY AANMAKEN ===" -ForegroundColor Cyan
    Write-Host ""
    
    $availabilityId = Read-InputWithValidation -Prompt "Geef een unieke ID voor de Availability" -Type "String" -Mandatory $true
    $name = Read-InputWithValidation -Prompt "Naam van de Availability" -Type "String" -Mandatory $true
    
    # Selecteer stores voor SalesPointGroup
    Write-Host "`nSelecteer store(s) voor deze Availability:" -ForegroundColor Yellow
    $stores = Get-CachedStores
    $selectedStore = Show-SelectionMenu -Title "Selecteer Store" -Items $stores -DisplayProperty "name" -IdProperty "id" -AllowCancel $false
    
    if (-not $selectedStore) {
        Write-Host "Geen store geselecteerd." -ForegroundColor Red
        Read-Host "`nDruk op Enter om door te gaan"
        return $null
    }
    
    # SalesPointGroup selecteren of aanmaken
    Write-Host "`nEen Availability heeft een SalesPointGroup nodig." -ForegroundColor Yellow
    $salesPointGroupId = Select-SalesPointGroup -Session $Session
    
    if (-not $salesPointGroupId) {
        Write-Host "Geen SalesPointGroup geselecteerd." -ForegroundColor Red
        Read-Host "`nDruk op Enter om door te gaan"
        return $null
    }
    
    # Selecteer TimePeriod
    Write-Host "`nSelecteer TimePeriod voor deze Availability:" -ForegroundColor Yellow
    $timePeriodId = Select-TimePeriod -Session $Session -Purpose "Selecteer TimePeriod voor Availability"
    
    if (-not $timePeriodId) {
        Write-Host "Geen TimePeriod geselecteerd." -ForegroundColor Red
        Read-Host "`nDruk op Enter om door te gaan"
        return $null
    }
    
    # FIX v1.3: Converteer IDs naar strings, trim spaties en valideer
    $salesPointGroupId = [string]$salesPointGroupId
    $salesPointGroupId = $salesPointGroupId.Trim()
    $timePeriodId = [string]$timePeriodId
    $timePeriodId = $timePeriodId.Trim()
    
    # Valideer dat beide IDs niet leeg zijn
    if ([string]::IsNullOrWhiteSpace($salesPointGroupId)) {
        Write-Host "`n[FOUT] SalesPointGroup ID is leeg!" -ForegroundColor Red
        Write-Log "FOUT: salesPointGroupId is leeg na selectie" -level ERROR
        Read-Host "`nDruk op Enter om door te gaan"
        return $null
    }
    
    if ([string]::IsNullOrWhiteSpace($timePeriodId)) {
        Write-Host "`n[FOUT] TimePeriod ID is leeg!" -ForegroundColor Red
        Write-Log "FOUT: timePeriodId is leeg na selectie" -level ERROR
        Read-Host "`nDruk op Enter om door te gaan"
        return $null
    }
    
    $body = @{
        name = $name
        salesPointGroup = $salesPointGroupId
        timePeriod = $timePeriodId
    }
    
    if ($DebugMode) {
        Write-Log "=== AVAILABILITY BODY ===" -level DEBUG
        Write-Log "salesPointGroup: '$salesPointGroupId'" -level DEBUG
        Write-Log "timePeriod: '$timePeriodId'" -level DEBUG
        Write-Log "name: '$name'" -level DEBUG
    }
    
    Write-Host "`nAvailability wordt aangemaakt..." -ForegroundColor Yellow
    
    try {
        Invoke-CMApi -Endpoint "availabilities/$availabilityId" -Method "PUT" -Body $body
        Write-Host "Availability succesvol aangemaakt!" -ForegroundColor Green
        
        Add-ToHistory -Session $Session -Action "Create" -EntityType "Availability" -EntityId $availabilityId -Parameters @{ Body = $body } -Result $null
        
        # FIX: Extra save om zeker te zijn dat availability wordt opgeslagen
        if (-not $Session.CreatedAvailabilities.Contains($availabilityId)) {
            $Session.CreatedAvailabilities += $availabilityId
        }
        # FIX v1.3: Suppress return waarde om array problemen te voorkomen
        [void](Save-Session -Session $Session)
        
        # Log voor verificatie
        if ($DebugMode) {
            Write-Log "Availability $availabilityId toegevoegd aan sessie. Totaal availabilities: $($Session.CreatedAvailabilities.Count)" -level DEBUG
        }
        
        Read-Host "`nDruk op Enter om door te gaan"
        return $availabilityId
    }
    catch {
        Write-Host "Fout bij aanmaken Availability: $_" -ForegroundColor Red
        Read-Host "`nDruk op Enter om door te gaan"
        return $null
    }
}

function Select-Availability {
    param(
        [PSCustomObject]$Session,
        [string]$Purpose = "Selecteer Availability",
        [bool]$AllowMultiple = $false
    )
    
    # Eerst kijken of er Availabilities in de sessie zijn
    $sessionAvailabilities = @()
    if ($Session.CreatedAvailabilities.Count -gt 0) {
        foreach ($avId in $Session.CreatedAvailabilities) {
            $sessionAvailabilities += [PSCustomObject]@{
                id = $avId
                name = "Availability uit sessie"
                source = "session"
            }
        }
    }
    
    Clear-Host
    Write-Host "`n=== $Purpose ===" -ForegroundColor Cyan
    Write-Host ""
    $avColor = if ($sessionAvailabilities.Count -gt 0) { "Green" } else { "Gray" }
    $avText = $("1. Availability uit deze sessie selecteren (" + $sessionAvailabilities.Count + " beschikbaar)")
    Write-Host $avText -ForegroundColor $avColor
    Write-Host "2. Nieuwe Availability aanmaken" -ForegroundColor Yellow
    Write-Host "3. Handmatig ID(s) invoeren" -ForegroundColor White
    Write-Host "C. Annuleren`n" -ForegroundColor Red
    
    $choice = Read-Host "Maak uw keuze"
    
    switch ($choice) {
        "1" {
            if ($sessionAvailabilities.Count -eq 0) {
                Write-Host "Geen Availabilities in deze sessie." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                return Select-Availability -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
            }
            
            $selected = Show-SelectionMenu -Title "Selecteer Availability uit sessie" -Items $sessionAvailabilities -DisplayProperty "id" -IdProperty "id" -AllowMultiple $AllowMultiple -AllowCancel $true
            if ($selected) {
                if ($AllowMultiple) {
                    # Zorg ervoor dat we een array van strings krijgen
                    $ids = @()
                    foreach ($item in $selected) {
                        if ($item.id) {
                            $ids += [string]$item.id
                        }
                    }
                    return $ids
                } else {
                    return [string]$selected.id
                }
            }
            return Select-Availability -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
        }
        "2" {
            $newId = New-Availability -Session $Session
            if ($newId) {
                if ($AllowMultiple) {
                    return @([string]$newId)
                } else {
                    return [string]$newId
                }
            }
            return Select-Availability -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
        }
        "3" {
            if ($AllowMultiple) {
                $manualIds = Read-InputWithValidation -Prompt "Geef Availability IDs (komma-gescheiden)" -Type "String" -Mandatory $true
                $ids = $manualIds -split ',' | ForEach-Object { [string]$_.Trim() }
                return $ids
            } else {
                $manualId = Read-InputWithValidation -Prompt "Geef Availability ID" -Type "String" -Mandatory $true
                return [string]$manualId
            }
        }
        "C" {
            return $null
        }
        "c" {
            return $null
        }
        default {
            return Select-Availability -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
        }
    }
}

function New-AvailabilityCore {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Session,

        [Parameter(Mandatory)]
        [string]$ExternalId,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        $SalesPointGroupId,

        [Parameter(Mandatory)]
        $TimePeriodId
    )

    # Normaliseer SalesPointGroupId -> string
    if ($SalesPointGroupId -is [string]) {
        $SalesPointGroupId = $SalesPointGroupId.Trim()
    } elseif ($null -ne $SalesPointGroupId -and $null -ne $SalesPointGroupId.PSObject.Properties['id']) {
        $SalesPointGroupId = ([string]$SalesPointGroupId.id).Trim()
    } else {
        throw "SalesPointGroupId is ongeldig (verwacht string of object met property 'id')."
    }

    # Normaliseer TimePeriodId -> string
    if ($TimePeriodId -is [string]) {
        $TimePeriodId = $TimePeriodId.Trim()
    } elseif ($null -ne $TimePeriodId -and $null -ne $TimePeriodId.PSObject.Properties['id']) {
        $TimePeriodId = ([string]$TimePeriodId.id).Trim()
    } else {
        throw "TimePeriodId is ongeldig (verwacht string of object met property 'id')."
    }

    if ([string]::IsNullOrWhiteSpace($SalesPointGroupId)) { throw "SalesPointGroupId is leeg." }
    if ([string]::IsNullOrWhiteSpace($TimePeriodId))     { throw "TimePeriodId is leeg." }

    $body = @{
        name            = $Name
        salesPointGroup = $SalesPointGroupId
        timePeriod      = $TimePeriodId
    }

    Invoke-CMApi -Endpoint "availabilities/$ExternalId" -Method "PUT" -Body $body

    Add-ToHistory -Session $Session -Action "Create" -EntityType "Availability" -EntityId $ExternalId -Parameters @{ Body = $body } -Result $null

    if (-not $Session.CreatedAvailabilities) { $Session.CreatedAvailabilities = @() }
    if ($Session.CreatedAvailabilities -notcontains $ExternalId) {
        $Session.CreatedAvailabilities += $ExternalId
    }

    [void](Save-Session -Session $Session)

    return $ExternalId
}



# ============================================================================
# CUSTOMER FUNCTIES
# ============================================================================


function Select-Customer {
    param(
        [PSCustomObject]$Session,
        [string]$Purpose = "Selecteer Customer",
        [bool]$AllowMultiple = $true
    )

    Clear-Host
    Write-Host "`n=== $Purpose ===" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "1. Customer(s) uit Onboarding API selecteren" -ForegroundColor Yellow
    Write-Host "2. Handmatig ID(s) invoeren" -ForegroundColor White
    Write-Host "3. Geen customers (overslaan)" -ForegroundColor Gray
    Write-Host "C. Annuleren`n" -ForegroundColor Red

    $choice = Read-Host "Maak uw keuze"

    switch ($choice) {
        "1" {
            $customers = Get-CachedCustomers  # moet onboarding-source zijn
            if (-not $customers -or $customers.Count -eq 0) {
                Write-Host "Geen customers gevonden in Onboarding API." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                return Select-Customer -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
            }

            Write-Host "Multi-select venster wordt geopend..." -ForegroundColor Yellow
            $pre = @()
            if ($Session.PSObject.Properties.Name -contains 'CreatedCustomers') {
                $pre = @($Session.CreatedCustomers)
            }

            $selectedIds = Show-CustomerMultiSelect -Customers $customers -PreSelectedIds $pre
            if ($selectedIds -and $selectedIds.Count -gt 0) {
                # Sla op in sessie (handig voor hergebruik, maar selectie komt altijd uit onboarding)
                if (-not ($Session.PSObject.Properties.Name -contains 'CreatedCustomers')) {
                    $Session | Add-Member -NotePropertyName "CreatedCustomers" -NotePropertyValue @() -Force
                }
                foreach ($id in $selectedIds) {
                    $sid = ([string]$id).Trim()
                    if ($sid -and $sid -notin $Session.CreatedCustomers) {
                        $Session.CreatedCustomers += $sid
                    }
                }

                if ($AllowMultiple) { return @($selectedIds) }
                return @([string]$selectedIds[0])
            }

            return Select-Customer -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
        }

        "2" {
            $manualIds = Read-InputWithValidation -Prompt "Geef Customer IDs (komma-gescheiden)" -Type "String" -Mandatory $true
            $ids = @($manualIds -split ',' | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })

            if (-not ($Session.PSObject.Properties.Name -contains 'CreatedCustomers')) {
                $Session | Add-Member -NotePropertyName "CreatedCustomers" -NotePropertyValue @() -Force
            }
            foreach ($id in $ids) {
                if ($id -notin $Session.CreatedCustomers) { $Session.CreatedCustomers += $id }
            }

            if ($AllowMultiple) { return $ids }
            return @($ids[0])
        }

        "3" { return @() }
        "C" { return $null }
        "c" { return $null }
        default { return Select-Customer -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple }
    }
}


function Select-CustomerGUI {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Session,

        [string]$Purpose = "Selecteer klanten",

        [array]$PreSelectedIds = @()
    )

    $customers = Get-CachedCustomers   # moet onboarding source zijn
    if (-not $customers -or $customers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Geen customers gevonden via Onboarding API.",
            "Customers",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return @()
    }

    $picked = Show-CustomerMultiSelect -Customers $customers -PreSelectedIds $PreSelectedIds
    if ($picked -and $picked.Count -gt 0) {
        return @($picked | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    }

    return @()
}



# Vervolg in Part 4...
# BooqCM-StamdataOnderhoud-Part4.ps1
# Dit bestand bevat de Promotion management logica

# ============================================================================
# PRODUCT SELECTIE
# ============================================================================

function Select-Product {
    param(
        [PSCustomObject]$Session,
        [string]$Purpose = "Selecteer Product",
        [bool]$AllowMultiple = $false
    )

    # Ensure CreatedProducts property bestaat
    if (-not ($Session.PSObject.Properties.Name -contains 'CreatedProducts')) {
        $Session | Add-Member -NotePropertyName "CreatedProducts" -NotePropertyValue @() -Force
    }

    # Eerst kijken of er Products in de sessie zijn (uit eerdere calls)
    $sessionProducts = @()
    if ($Session.CreatedProducts -and $Session.CreatedProducts.Count -gt 0) {
        foreach ($prodId in $Session.CreatedProducts) {
            $sessionProducts += [PSCustomObject]@{
                id     = [string]$prodId
                name   = "Product uit sessie"
                source = "session"
            }
        }
    }

    Clear-Host
    Write-Host "`n=== $Purpose ===" -ForegroundColor Cyan
    Write-Host ""
    $prodColor = if ($sessionProducts.Count -gt 0) { "Green" } else { "Gray" }
    $prodText = $("1. Product uit deze sessie selecteren (" + $sessionProducts.Count + " beschikbaar)")
    Write-Host $prodText -ForegroundColor $prodColor
    Write-Host "2. Product uit PIM API selecteren" -ForegroundColor Yellow
    Write-Host "3. Handmatig ID invoeren" -ForegroundColor White
    Write-Host "C. Annuleren`n" -ForegroundColor Red

    $choice = Read-Host "Maak uw keuze"

    switch ($choice) {
        "1" {
            if ($sessionProducts.Count -eq 0) {
                Write-Host "Geen Products in deze sessie." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                return Select-Product -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
            }

            $selected = Show-SelectionMenu `
                -Title "Selecteer Product uit sessie" `
                -Items $sessionProducts `
                -DisplayProperty "id" `
                -IdProperty "id" `
                -AllowMultiple $AllowMultiple `
                -AllowCancel $true

            if ($selected) {
                if ($AllowMultiple) {
                    # Normalizeer naar string[]
                    $ids = @()
                    foreach ($item in @($selected)) {
                        if ($null -ne $item.PSObject.Properties['id'] -and -not [string]::IsNullOrWhiteSpace($item.id)) {
                            $ids += ([string]$item.id).Trim()
                        } elseif ($item -is [string] -and -not [string]::IsNullOrWhiteSpace($item)) {
                            $ids += ([string]$item).Trim()
                        }
                    }
                    return $ids
                } else {
                    return ([string]$selected.id).Trim()
                }
            }

            return Select-Product -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
        }

        "2" {
            $products = Get-CachedProducts
            if ($products.Count -eq 0) {
                Write-Host "Geen products gevonden in PIM API." -ForegroundColor Yellow
                Start-Sleep -Seconds 2
                return Select-Product -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
            }

            # Voor grote lijsten (>20), gebruik altijd Windows Form multi-select
            if ($products.Count -gt 20) {
                Write-Host "`nEr zijn $($products.Count) producten beschikbaar." -ForegroundColor Cyan
                Write-Host "Multi-select venster wordt geopend..." -ForegroundColor Yellow

                $preSelected = if ($Session.CreatedProducts) { $Session.CreatedProducts } else { @() }
                $selectedIds = Show-ProductMultiSelect -Products $products -PreSelectedIds $preSelected

                if ($selectedIds -and $selectedIds.Count -gt 0) {
                    if ($DebugMode) {
                        Write-Log "Product multi-select resultaat: $($selectedIds.Count) producten geselecteerd" -level DEBUG
                    }

                    foreach ($id in $selectedIds) {
                        if ($id -and $id -notin $Session.CreatedProducts) {
                            $Session.CreatedProducts += ([string]$id).Trim()
                        }
                    }

                    if (-not $AllowMultiple) {
                        Write-Host "Geselecteerd: $($selectedIds[0])" -ForegroundColor Green
                        return ([string]$selectedIds[0]).Trim()
                    }

                    Write-Host "Geselecteerd: $($selectedIds.Count) producten" -ForegroundColor Green
                    return @($selectedIds | ForEach-Object { ([string]$_).Trim() })
                }

                return Select-Product -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
            }
            else {
                # Voor kleine lijsten, gebruik oude methode met zoeken
                Write-Host "`nZoeken in producten (leeg laten voor alle producten tonen):" -ForegroundColor Yellow
                $searchTerm = Read-Host "Zoekterm"

                if (-not [string]::IsNullOrWhiteSpace($searchTerm)) {
                    $products = $products | Where-Object { $_.name -like "*$searchTerm*" -or $_.id -like "*$searchTerm*" }
                    Write-Host "Gevonden: $($products.Count) producten" -ForegroundColor Cyan
                }

                if ($products.Count -eq 0) {
                    Write-Host "Geen producten gevonden met deze zoekterm." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                    return Select-Product -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
                }

                $selected = Show-SelectionMenu `
                    -Title "Selecteer Product uit PIM API" `
                    -Items $products `
                    -DisplayProperty "name" `
                    -IdProperty "id" `
                    -AllowCancel $true

                if ($selected) {
                    if ($selected.id -notin $Session.CreatedProducts) {
                        $Session.CreatedProducts += ([string]$selected.id).Trim()
                    }
                    return ([string]$selected.id).Trim()
                }

                return Select-Product -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
            }
        }

        "3" {
            if ($AllowMultiple) {
                $manualIds = Read-InputWithValidation -Prompt "Geef Product IDs (komma-gescheiden)" -Type "String" -Mandatory $true
                $ids = @($manualIds -split ',' | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })

                foreach ($id in $ids) {
                    if ($id -notin $Session.CreatedProducts) {
                        $Session.CreatedProducts += $id
                    }
                }
                return $ids
            } else {
                $manualId = Read-InputWithValidation -Prompt "Geef Product ID" -Type "String" -Mandatory $true
                $manualId = ([string]$manualId).Trim()

                if ($manualId -notin $Session.CreatedProducts) {
                    $Session.CreatedProducts += $manualId
                }
                return $manualId
            }
        }

        "C" { return $null }
        "c" { return $null }

        default {
            return Select-Product -Session $Session -Purpose $Purpose -AllowMultiple $AllowMultiple
        }
    }
}

# ============================================================================
# PROMOTION FUNCTIES
# ============================================================================




function Add-PromotionToProduct {
    param([PSCustomObject]$Session)
    
    Clear-Host
    Write-Host "`n+==============================================================" -ForegroundColor Cyan
    Write-Host "|         PROMOTIE TOEVOEGEN AAN PRODUCT                       |" -ForegroundColor Cyan
    Write-Host "+==============================================================+" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Deze wizard helpt u om een promotie aan een product toe te voegen." -ForegroundColor Yellow
    Write-Host "U moet de volgende stappen doorlopen:" -ForegroundColor Yellow
    Write-Host "  1. Product selecteren" -ForegroundColor Gray
    Write-Host "  2. Promotie aanmaken of selecteren" -ForegroundColor Gray
    Write-Host "  3. TimePeriod(s) aanmaken indien nodig" -ForegroundColor Gray
    Write-Host "  4. Availability(s) aanmaken indien nodig" -ForegroundColor Gray
    Write-Host "  5. Customer selectie (optioneel)" -ForegroundColor Gray
    Write-Host ""
    
    Read-Host "Druk op Enter om te starten"
    
    # Start met het aanmaken van de promotie
    # Dit proces zal alle benodigde stappen doorlopen
    $promotionId = New-Promotion -Session $Session
    
    if ($promotionId) {
        Write-Host "`n[OK] Promotie succesvol aangemaakt en gekoppeld!" -ForegroundColor Green
        Write-Host "Promotie ID: $promotionId" -ForegroundColor Cyan
    }
    else {
        Write-Host "`n[FOUT] Promotie niet aangemaakt." -ForegroundColor Red
    }
    
    Read-Host "`nDruk op Enter om terug te gaan naar het hoofdmenu"
}


function New-PromotionCore {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Session,

        [Parameter(Mandatory)]
        [string]$ExternalId,

        [Parameter(Mandatory)]
        [string]$Name,

        [Parameter(Mandatory)]
        [ValidateSet("TICKET_DISCOUNT","QUANTITY_DISCOUNT","COMBI_DEAL")]
        [string]$Type,

        [Parameter(Mandatory)]
        [ValidateSet("NEW_AMOUNT","DISCOUNT_AMOUNT","DISCOUNT_PERCENTAGE")]
        [string]$EffectType,

        [Parameter(Mandatory)]
        [decimal]$EffectValue,

        # YAML ProductCondition supports multiple products
        [string[]]$ProductIds,

        # backwards compatible: single product id
        [string]$ProductId,

        # YAML uses 'quantity' (required)
        [int]$Quantity = 1,

        # YAML TicketCondition.minimumAmount is required
        [decimal]$TicketMinimumAmount,

        [string[]]$AvailabilityIds,

        [string[]]$CustomerIds,

        # YAML: names: object<string,string>
        [hashtable]$Names
    )

    $ExternalId = $ExternalId.Trim()
    $Name = $Name.Trim()

    if ([string]::IsNullOrWhiteSpace($ExternalId)) { throw "ExternalId is verplicht." }
    if ([string]::IsNullOrWhiteSpace($Name)) { throw "Naam is verplicht." }

    $effect = @{
        type  = $EffectType
        value = [decimal]$EffectValue
    }

    $body = @{
        name   = $Name
        type   = $Type
        effect = $effect
    }

    # ---- names translations (YAML-conform) ----
    if ($PSBoundParameters.ContainsKey('Names') -and $Names) {
        # normalize: remove null keys, keep strings (empty string allowed per spec)
        $clean = @{}
        foreach ($k in $Names.Keys) {
            if ([string]::IsNullOrWhiteSpace([string]$k)) { continue }
            $clean[[string]$k] = [string]$Names[$k]
        }
        if ($clean.Keys.Count -gt 0) {
            $body['names'] = $clean
        }
    }

    # ---- productConditions (YAML-conform) ----
    $normalizedProductIds = @()
    if ($PSBoundParameters.ContainsKey('ProductIds') -and $ProductIds) {
        $normalizedProductIds = @($ProductIds | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
    } elseif ($PSBoundParameters.ContainsKey('ProductId') -and -not [string]::IsNullOrWhiteSpace($ProductId)) {
        $normalizedProductIds = @(([string]$ProductId).Trim())
    }

    if ($Type -in @("QUANTITY_DISCOUNT","COMBI_DEAL")) {
        if (-not $normalizedProductIds -or $normalizedProductIds.Count -eq 0) {
            throw "$Type vereist minimaal 1 product."
        }
        if ($Type -eq "COMBI_DEAL" -and $normalizedProductIds.Count -lt 2) {
            throw "COMBI_DEAL vereist minimaal 2 producten."
        }
        if ($Quantity -lt 1) { throw "Quantity moet minimaal 1 zijn." }

        $body['productConditions'] = @(
            @{
                products = $normalizedProductIds
                quantity = [int]$Quantity
            }
        )
    }

    # ---- ticketCondition (YAML-conform, minimumAmount required) ----
    if ($Type -eq "TICKET_DISCOUNT") {
        if (-not $PSBoundParameters.ContainsKey('TicketMinimumAmount')) {
            throw "TICKET_DISCOUNT vereist ticketCondition.minimumAmount."
        }
        $body['ticketCondition'] = @{ minimumAmount = [decimal]$TicketMinimumAmount }
    }

    # ---- availability ----
    if ($AvailabilityIds -and $AvailabilityIds.Count -gt 0) {
        $body['availability'] = @($AvailabilityIds | ForEach-Object { ([string]$_.Trim()) } | Where-Object { $_ })
    }

    # ---- customerCondition (YAML-conform) ----
    if ($CustomerIds -and $CustomerIds.Count -gt 0) {
        $body['customerCondition'] = @{
            customers = @($CustomerIds | ForEach-Object { ([string]$_.Trim()) } | Where-Object { $_ })
        }
    }

    $result = Invoke-CMApi -Endpoint "promotions/$ExternalId" -Method "PUT" -Body $body
    Add-ToHistory -Session $Session -Action "Create" -EntityType "Promotion" -EntityId $ExternalId -Parameters @{ Body = $body } -Result $result

    if (-not $Session.CreatedPromotions) { $Session.CreatedPromotions = @() }
    if ($Session.CreatedPromotions -notcontains $ExternalId) { $Session.CreatedPromotions += $ExternalId }

    [void](Save-Session -Session $Session)

    return $ExternalId
}




function Show-NewPromotionGUI {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Session,
        [System.Windows.Forms.Form]$OwnerForm
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Nieuwe promotie aanmaken"
    $dlg.StartPosition = "CenterParent"
    $dlg.Size = New-Object System.Drawing.Size(980, 640)
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false

    # Centrale state
    $dlg.Tag = @{
        ProductIds  = @()
        CustomerIds = @()
    }

    function Refresh-Availabilities {
        $lstAvail.Items.Clear()
        $av = @()
        if ($Session.CreatedAvailabilities) { $av = @($Session.CreatedAvailabilities) }
        foreach ($id in $av) { [void]$lstAvail.Items.Add([string]$id) }
    }

    # ----------------------------
    # ID / Name
    # ----------------------------
    $lblId = New-Object System.Windows.Forms.Label
    $lblId.Location = New-Object System.Drawing.Point(16, 16)
    $lblId.Size = New-Object System.Drawing.Size(120, 20)
    $lblId.Text = "ExternalId:"
    $dlg.Controls.Add($lblId)

    $txtId = New-Object System.Windows.Forms.TextBox
    $txtId.Location = New-Object System.Drawing.Point(150, 14)
    $txtId.Size = New-Object System.Drawing.Size(800, 20)
    $dlg.Controls.Add($txtId)

    $lblName = New-Object System.Windows.Forms.Label
    $lblName.Location = New-Object System.Drawing.Point(16, 48)
    $lblName.Size = New-Object System.Drawing.Size(120, 20)
    $lblName.Text = "Naam:"
    $dlg.Controls.Add($lblName)

    $txtName = New-Object System.Windows.Forms.TextBox
    $txtName.Location = New-Object System.Drawing.Point(150, 46)
    $txtName.Size = New-Object System.Drawing.Size(800, 20)
    $dlg.Controls.Add($txtName)

    # ----------------------------
    # Type / Effect
    # ----------------------------
    $lblType = New-Object System.Windows.Forms.Label
    $lblType.Location = New-Object System.Drawing.Point(16, 84)
    $lblType.Size = New-Object System.Drawing.Size(120, 20)
    $lblType.Text = "Type:"
    $dlg.Controls.Add($lblType)

    $cmbType = New-Object System.Windows.Forms.ComboBox
    $cmbType.Location = New-Object System.Drawing.Point(150, 82)
    $cmbType.Size = New-Object System.Drawing.Size(220, 20)
    $cmbType.DropDownStyle = "DropDownList"
    @("TICKET_DISCOUNT","QUANTITY_DISCOUNT","COMBI_DEAL") | ForEach-Object { [void]$cmbType.Items.Add($_) }
    $cmbType.SelectedItem = "TICKET_DISCOUNT"
    $dlg.Controls.Add($cmbType)

    $lblEff = New-Object System.Windows.Forms.Label
    $lblEff.Location = New-Object System.Drawing.Point(16, 116)
    $lblEff.Size = New-Object System.Drawing.Size(120, 20)
    $lblEff.Text = "Effect type:"
    $dlg.Controls.Add($lblEff)

    $cmbEff = New-Object System.Windows.Forms.ComboBox
    $cmbEff.Location = New-Object System.Drawing.Point(150, 114)
    $cmbEff.Size = New-Object System.Drawing.Size(220, 20)
    $cmbEff.DropDownStyle = "DropDownList"
    @("DISCOUNT_PERCENTAGE","DISCOUNT_AMOUNT","NEW_AMOUNT") | ForEach-Object { [void]$cmbEff.Items.Add($_) }
    $cmbEff.SelectedItem = "DISCOUNT_PERCENTAGE"
    $dlg.Controls.Add($cmbEff)

    $lblVal = New-Object System.Windows.Forms.Label
    $lblVal.Location = New-Object System.Drawing.Point(390, 116)
    $lblVal.Size = New-Object System.Drawing.Size(120, 20)
    $lblVal.Text = "Waarde:"
    $dlg.Controls.Add($lblVal)

    $numVal = New-Object System.Windows.Forms.NumericUpDown
    $numVal.Location = New-Object System.Drawing.Point(($lblVal.Right + 8), 114)
    $numVal.Size = New-Object System.Drawing.Size(120, 20)
    $numVal.DecimalPlaces = 2
    $numVal.Maximum = 1000000
    $numVal.Value = 5
    $dlg.Controls.Add($numVal)

    # ----------------------------
    # TicketCondition (checkbox)
    # ----------------------------
    $chkTicket = New-Object System.Windows.Forms.CheckBox
    $chkTicket.Location = New-Object System.Drawing.Point(150, 150)
    $chkTicket.Size = New-Object System.Drawing.Size(240, 20)
    $chkTicket.Text = "Ticket condition opnemen"
    $dlg.Controls.Add($chkTicket)

    $lblMinAmt = New-Object System.Windows.Forms.Label
    $lblMinAmt.Location = New-Object System.Drawing.Point(390, 150)
    $lblMinAmt.Size = New-Object System.Drawing.Size(180, 20)
    $lblMinAmt.Text = "MinimumAmount:"
    $dlg.Controls.Add($lblMinAmt)

    $numMin = New-Object System.Windows.Forms.NumericUpDown
    $numMin.Location = New-Object System.Drawing.Point(580, 148)
    $numMin.Size = New-Object System.Drawing.Size(120, 20)
    $numMin.DecimalPlaces = 2
    $numMin.Maximum = 1000000
    $numMin.Value = 1
    $dlg.Controls.Add($numMin)

    # ----------------------------
    # ProductConditions (checkbox)
    # YAML: products[] + quantity (required)
    # ----------------------------
    $chkProdCond = New-Object System.Windows.Forms.CheckBox
    $chkProdCond.Location = New-Object System.Drawing.Point(150, 182)
    $chkProdCond.Size = New-Object System.Drawing.Size(240, 20)
    $chkProdCond.Text = "Product conditions opnemen"
    $dlg.Controls.Add($chkProdCond)

    $lblProd = New-Object System.Windows.Forms.Label
    $lblProd.Location = New-Object System.Drawing.Point(16, 214)
    $lblProd.Size = New-Object System.Drawing.Size(120, 20)
    $lblProd.Text = "Products:"
    $dlg.Controls.Add($lblProd)

    $txtProd = New-Object System.Windows.Forms.TextBox
    $txtProd.Location = New-Object System.Drawing.Point(150, 212)
    $txtProd.Size = New-Object System.Drawing.Size(520, 20)
    $txtProd.ReadOnly = $true
    $dlg.Controls.Add($txtProd)

    $btnPickProd = New-Object System.Windows.Forms.Button
    $btnPickProd.Location = New-Object System.Drawing.Point(680, 208)
    $btnPickProd.Size = New-Object System.Drawing.Size(270, 28)
    $btnPickProd.Text = "Product(en) kiezen..."
    $dlg.Controls.Add($btnPickProd)

    $lblQty = New-Object System.Windows.Forms.Label
    $lblQty.Location = New-Object System.Drawing.Point(16, 246)
    $lblQty.Size = New-Object System.Drawing.Size(120, 20)
    $lblQty.Text = "Quantity:"
    $dlg.Controls.Add($lblQty)

    $numQty = New-Object System.Windows.Forms.NumericUpDown
    $numQty.Location = New-Object System.Drawing.Point(150, 244)
    $numQty.Size = New-Object System.Drawing.Size(120, 20)
    $numQty.Minimum = 1
    $numQty.Maximum = 9999
    $numQty.Value = 1
    $dlg.Controls.Add($numQty)

    $btnPickProd.Add_Click({
        try {
            $products = Get-CachedProducts
            $picked = Show-ProductMultiSelect -Products $products -PreSelectedIds @($dlg.Tag.ProductIds)
            if ($picked -and $picked.Count -gt 0) {
                $dlg.Tag.ProductIds = @($picked | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ })
                $txtProd.Text = ($dlg.Tag.ProductIds -join ", ")
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fout bij productselectie: $($_.Exception.Message)") | Out-Null
        }
    })

    # ----------------------------
    # Availability (checkbox)
    # ----------------------------
    $chkAvail = New-Object System.Windows.Forms.CheckBox
    $chkAvail.Location = New-Object System.Drawing.Point(16, 286)
    $chkAvail.Size = New-Object System.Drawing.Size(280, 20)
    $chkAvail.Text = "Availability opnemen in request"
    $chkAvail.Checked = $true
    $dlg.Controls.Add($chkAvail)

    $grpAv = New-Object System.Windows.Forms.GroupBox
    $grpAv.Location = New-Object System.Drawing.Point(16, 310)
    $grpAv.Size = New-Object System.Drawing.Size(450, 240)
    $grpAv.Text = "Availabilities (uit sessie)"
    $dlg.Controls.Add($grpAv)

    $lstAvail = New-Object System.Windows.Forms.ListBox
    $lstAvail.Location = New-Object System.Drawing.Point(12, 24)
    $lstAvail.Size = New-Object System.Drawing.Size(426, 160)
    $lstAvail.SelectionMode = "MultiExtended"
    $grpAv.Controls.Add($lstAvail)

    $btnNewAv = New-Object System.Windows.Forms.Button
    $btnNewAv.Location = New-Object System.Drawing.Point(12, 192)
    $btnNewAv.Size = New-Object System.Drawing.Size(426, 28)
    $btnNewAv.Text = "Nieuwe Availability aanmaken..."
    $grpAv.Controls.Add($btnNewAv)

    $btnNewAv.Add_Click({
        try {
            Show-NewAvailabilityGUI -Session $Session -OwnerForm $dlg
            Refresh-Availabilities
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Fout bij Availability: $($_.Exception.Message)") | Out-Null
        }
    })

    # ----------------------------
    # CustomerCondition (checkbox + selector)
    # YAML: customerCondition: { customers: string[] }
    # ----------------------------
    $chkCust = New-Object System.Windows.Forms.CheckBox
    $chkCust.Location = New-Object System.Drawing.Point(490, 310)
    $chkCust.Size = New-Object System.Drawing.Size(320, 20)
    $chkCust.Text = "Customer condition opnemen"
    $dlg.Controls.Add($chkCust)

    $lblCust = New-Object System.Windows.Forms.Label
    $lblCust.Location = New-Object System.Drawing.Point(490, 336)
    $lblCust.Size = New-Object System.Drawing.Size(460, 20)
    $lblCust.Text = "Customers (IDs):"
    $dlg.Controls.Add($lblCust)

    $txtCust = New-Object System.Windows.Forms.TextBox
    $txtCust.Location = New-Object System.Drawing.Point(490, 360)
    $txtCust.Size = New-Object System.Drawing.Size(460, 60)
    $txtCust.Multiline = $true
    $dlg.Controls.Add($txtCust)

    $btnPickCust = New-Object System.Windows.Forms.Button
    $btnPickCust.Location = New-Object System.Drawing.Point(490, 426)
    $btnPickCust.Size = New-Object System.Drawing.Size(460, 28)
    $btnPickCust.Text = "Customers selecteren..."
    $dlg.Controls.Add($btnPickCust)

	$btnPickCust.Add_Click({
		try {
			$pre = @()
			if ($dlg.Tag -and $dlg.Tag.CustomerIds) { $pre = @($dlg.Tag.CustomerIds) }
	
			$picked = Select-CustomerGUI -Session $Session -Purpose "Selecteer klanten" -PreSelectedIds $pre
			if ($picked -and $picked.Count -gt 0) {
				$dlg.Tag.CustomerIds = @($picked)
				$txtCust.Text = ($dlg.Tag.CustomerIds -join ", ")
				$chkCust.Checked = $true
			}
		} catch {
			[System.Windows.Forms.MessageBox]::Show("Fout bij customer selectie: $($_.Exception.Message)") | Out-Null
		}
	})
	
	

    # ----------------------------
    # Names (translations) (checkbox + EN/NL)
    # YAML: names: object { [lang]: string }
    # ----------------------------
    $chkNames = New-Object System.Windows.Forms.CheckBox
    $chkNames.Location = New-Object System.Drawing.Point(490, 466)
    $chkNames.Size = New-Object System.Drawing.Size(320, 20)
    $chkNames.Text = "Translations (names) opnemen"
    $dlg.Controls.Add($chkNames)

    $lblEN = New-Object System.Windows.Forms.Label
    $lblEN.Location = New-Object System.Drawing.Point(490, 490)
    $lblEN.Size = New-Object System.Drawing.Size(40, 20)
    $lblEN.Text = "EN:"
    $dlg.Controls.Add($lblEN)

    $txtEN = New-Object System.Windows.Forms.TextBox
    $txtEN.Location = New-Object System.Drawing.Point(535, 488)
    $txtEN.Size = New-Object System.Drawing.Size(415, 20)
    $dlg.Controls.Add($txtEN)

    $lblNL = New-Object System.Windows.Forms.Label
    $lblNL.Location = New-Object System.Drawing.Point(490, 520)
    $lblNL.Size = New-Object System.Drawing.Size(40, 20)
    $lblNL.Text = "NL:"
    $dlg.Controls.Add($lblNL)

    $txtNL = New-Object System.Windows.Forms.TextBox
    $txtNL.Location = New-Object System.Drawing.Point(535, 518)
    $txtNL.Size = New-Object System.Drawing.Size(415, 20)
    $dlg.Controls.Add($txtNL)

    # enable/disable blocks based on checkboxes and type
    $applyRules = {
        $t = [string]$cmbType.SelectedItem

        # ticketCondition only relevant for ticket discount
        $ticketType = ($t -eq "TICKET_DISCOUNT")
        $chkTicket.Enabled = $ticketType
        if ($ticketType) {
            $chkTicket.Checked = $true
        } else {
            $chkTicket.Checked = $false
        }
        $numMin.Enabled = $chkTicket.Checked

        # productConditions only relevant for quantity/combi
        $prodType = ($t -in @("QUANTITY_DISCOUNT","COMBI_DEAL"))
        $chkProdCond.Enabled = $prodType
        if ($prodType) {
            $chkProdCond.Checked = $true
        } else {
            $chkProdCond.Checked = $false
            $dlg.Tag.ProductIds = @()
            $txtProd.Text = ""
        }

        $btnPickProd.Enabled = $chkProdCond.Checked
        $numQty.Enabled = $chkProdCond.Checked
    }

    $chkTicket.Add_CheckedChanged({ $numMin.Enabled = $chkTicket.Checked })
    $chkProdCond.Add_CheckedChanged({
        $btnPickProd.Enabled = $chkProdCond.Checked
        $numQty.Enabled = $chkProdCond.Checked
        if (-not $chkProdCond.Checked) {
            $dlg.Tag.ProductIds = @()
            $txtProd.Text = ""
        }
    })

    $chkCust.Add_CheckedChanged({
        $txtCust.Enabled = $chkCust.Checked
        $btnPickCust.Enabled = $chkCust.Checked
        if (-not $chkCust.Checked) {
            $dlg.Tag.CustomerIds = @()
            $txtCust.Text = ""
        }
    })
    $chkCust.Checked = $false
    $txtCust.Enabled = $false
    $btnPickCust.Enabled = $false

    $chkNames.Add_CheckedChanged({
        $txtEN.Enabled = $chkNames.Checked
        $txtNL.Enabled = $chkNames.Checked
        if (-not $chkNames.Checked) {
            $txtEN.Text = ""
            $txtNL.Text = ""
        }
    })
    $chkNames.Checked = $false
    $txtEN.Enabled = $false
    $txtNL.Enabled = $false

    $chkAvail.Add_CheckedChanged({
        $lstAvail.Enabled = $chkAvail.Checked
        $btnNewAv.Enabled = $chkAvail.Checked
        if (-not $chkAvail.Checked) {
            $lstAvail.ClearSelected()
        }
    })

    $cmbType.Add_SelectedIndexChanged($applyRules)
    & $applyRules

    # ----------------------------
    # Buttons
    # ----------------------------
    $chkKeepOpen = New-Object System.Windows.Forms.CheckBox
    $chkKeepOpen.Location = New-Object System.Drawing.Point(16, 560)
    $chkKeepOpen.Size = New-Object System.Drawing.Size(260, 20)
    $chkKeepOpen.Text = "Open houden na Opslaan"
    $chkKeepOpen.Checked = $true
    $dlg.Controls.Add($chkKeepOpen)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Location = New-Object System.Drawing.Point(754, 556)
    $btnOk.Size = New-Object System.Drawing.Size(96, 32)
    $btnOk.Text = "Opslaan"
    $dlg.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Location = New-Object System.Drawing.Point(854, 556)
    $btnCancel.Size = New-Object System.Drawing.Size(96, 32)
    $btnCancel.Text = "Sluiten"
    $btnCancel.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($btnCancel)

    $btnOk.Add_Click({
        try {
            $externalId = $txtId.Text.Trim()
            $name = $txtName.Text.Trim()
            if ([string]::IsNullOrWhiteSpace($externalId)) { throw "ExternalId is verplicht." }
            if ([string]::IsNullOrWhiteSpace($name)) { throw "Naam is verplicht." }

            $type = [string]$cmbType.SelectedItem
            $effType = [string]$cmbEff.SelectedItem
            $effVal = [decimal]$numVal.Value

            $params = @{
                Session        = $Session
                ExternalId     = $externalId
                Name           = $name
                Type           = $type
                EffectType     = $effType
                EffectValue    = $effVal
            }

            # Availability (optioneel via checkbox)
            if ($chkAvail.Checked) {
                $avIds = @()
                foreach ($sel in $lstAvail.SelectedItems) { $avIds += ([string]$sel).Trim() }
                if ($avIds.Count -eq 0) { throw "Availability opnemen is aangevinkt, maar er is niets geselecteerd." }
                $params.AvailabilityIds = $avIds
            }

            # ProductConditions (optioneel via checkbox, maar forced per type)
            if ($type -in @("QUANTITY_DISCOUNT","COMBI_DEAL")) {
                if (-not $chkProdCond.Checked) { throw "Product conditions zijn verplicht voor dit type (YAML)." }
                $pids = @($dlg.Tag.ProductIds)
                if (-not $pids -or $pids.Count -eq 0) { throw "Selecteer minimaal 1 product." }
                $params.ProductIds = $pids
                $params.Quantity = [int]$numQty.Value
            }

            # TicketCondition (forced per type, YAML minimumAmount required)
            if ($type -eq "TICKET_DISCOUNT") {
                if (-not $chkTicket.Checked) { throw "Ticket condition is verplicht voor TICKET_DISCOUNT (YAML)." }
                $params.TicketMinimumAmount = [decimal]$numMin.Value
            }

            # CustomerCondition
            if ($chkCust.Checked) {
                $custIds = @()
                if (-not [string]::IsNullOrWhiteSpace($txtCust.Text)) {
                    $custIds = @($txtCust.Text -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                }
                if (-not $custIds -or $custIds.Count -eq 0) { throw "Customer condition opnemen is aangevinkt, maar er zijn geen customer IDs." }
                $params.CustomerIds = $custIds
            }


			# Names translations
			if ($chkNames.Checked) {
				$names = @{}
				if (-not [string]::IsNullOrWhiteSpace($txtEN.Text)) { $names['EN'] = $txtEN.Text }
				if (-not [string]::IsNullOrWhiteSpace($txtNL.Text)) { $names['NL'] = $txtNL.Text }
			
				if ($names.Keys.Count -eq 0) {
					throw "Translations (names) is aangevinkt maar EN/NL is leeg."
				}
			
				$params.Names = $names
			}
			
			
            $btnOk.Enabled = $false
            $dlg.UseWaitCursor = $true

            $id = New-PromotionCore @params

            [System.Windows.Forms.MessageBox]::Show("Promotie opgeslagen: $id", "Succes") | Out-Null
            if (-not $chkKeepOpen.Checked) {
                $dlg.Close()
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fout: $($_.Exception.Message)", "Fout") | Out-Null
        }
        finally {
            $dlg.UseWaitCursor = $false
            $btnOk.Enabled = $true
        }
    })

    Refresh-Availabilities
    $null = $dlg.ShowDialog($OwnerForm)
}





# Vervolg in Part 5 met het hoofdmenu...
# BooqCM-StamdataOnderhoud-Part5.ps1
# Dit bestand bevat het hoofdmenu en main loop

# ============================================================================
# HOOFDMENU
# ============================================================================

function Show-MainMenuGUI {
    param([PSCustomObject]$Session)
    
    # FIX v1.3: Forms assembly loading met error handling
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
    }
    catch {
        Write-Log "FOUT: Windows Forms kon niet worden geladen: $_" -level ERROR
        Write-Host "[FOUT] GUI kan niet worden gestart. Forms assembly niet beschikbaar." -ForegroundColor Red
        throw "Windows Forms not available"
    }
    
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Booq CM Stamdata Onderhoud"
    $form.Size = New-Object System.Drawing.Size(600, 680)
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "FixedDialog"
    $form.MaximizeBox = $false
    
    # Sessie info panel
    $infoPanel = New-Object System.Windows.Forms.Panel
    $infoPanel.Location = New-Object System.Drawing.Point(10, 10)
    $infoPanel.Size = New-Object System.Drawing.Size(560, 80)
    $infoPanel.BorderStyle = "FixedSingle"
    $form.Controls.Add($infoPanel)
    
    $lblSessionName = New-Object System.Windows.Forms.Label
    $lblSessionName.Location = New-Object System.Drawing.Point(10, 10)
    $lblSessionName.Size = New-Object System.Drawing.Size(540, 20)
    $lblSessionName.Text = "Sessie: $($Session.Name)"
    $lblSessionName.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $infoPanel.Controls.Add($lblSessionName)
    
    $lblEnv = New-Object System.Windows.Forms.Label
    $lblEnv.Location = New-Object System.Drawing.Point(10, 35)
    $lblEnv.Size = New-Object System.Drawing.Size(250, 20)
    $lblEnv.Text = "Omgeving: $environment"
    $infoPanel.Controls.Add($lblEnv)
    
    $lblActions = New-Object System.Windows.Forms.Label
    $lblActions.Location = New-Object System.Drawing.Point(10, 55)
    $lblActions.Size = New-Object System.Drawing.Size(250, 20)
    $lblActions.Text = "Aantal acties: $($Session.History.Count)"
    $infoPanel.Controls.Add($lblActions)
    
    $lblStats = New-Object System.Windows.Forms.Label
    $lblStats.Location = New-Object System.Drawing.Point(280, 35)
    $lblStats.Size = New-Object System.Drawing.Size(270, 40)
    $tpCount = if ($Session.CreatedTimePeriods) { $Session.CreatedTimePeriods.Count } else { 0 }
    $avCount = if ($Session.CreatedAvailabilities) { $Session.CreatedAvailabilities.Count } else { 0 }
    $promCount = if ($Session.CreatedPromotions) { $Session.CreatedPromotions.Count } else { 0 }
    $lblStats.Text = "TP: $tpCount | Avail: $avCount | Prom: $promCount"
    $infoPanel.Controls.Add($lblStats)
    
	# Promoties sectie
	$lblPromotions = New-Object System.Windows.Forms.Label
	$lblPromotions.Location = New-Object System.Drawing.Point(20, 105)
	$lblPromotions.Size = New-Object System.Drawing.Size(560, 25)
	$lblPromotions.Text = "> PROMOTIES"
	$lblPromotions.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
	$lblPromotions.ForeColor = [System.Drawing.Color]::DarkBlue
	$form.Controls.Add($lblPromotions)
	
	$btnPromotion = New-Object System.Windows.Forms.Button
	$btnPromotion.Location = New-Object System.Drawing.Point(30, 135)
	$btnPromotion.Size = New-Object System.Drawing.Size(540, 35)
	$btnPromotion.Text = "Promotie aanmaken"
	$btnPromotion.BackColor = [System.Drawing.Color]::LightGreen
	$btnPromotion.Add_Click({
		Show-NewPromotionGUI -Session $Session -OwnerForm $form
	})
	$form.Controls.Add($btnPromotion)
	
	# Verwijder/laat weg:
	# - $btnAddPromotion
	# - $btnNewPromotion
   
    # Ondersteunende entiteiten
    $lblSupport = New-Object System.Windows.Forms.Label
    $lblSupport.Location = New-Object System.Drawing.Point(20, 220)
    $lblSupport.Size = New-Object System.Drawing.Size(560, 25)
    $lblSupport.Text = "> ONDERSTEUNENDE ENTITEITEN"
    $lblSupport.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $lblSupport.ForeColor = [System.Drawing.Color]::DarkBlue
    $form.Controls.Add($lblSupport)
    
    $btnTimePeriod = New-Object System.Windows.Forms.Button
    $btnTimePeriod.Location = New-Object System.Drawing.Point(30, 250)
    $btnTimePeriod.Size = New-Object System.Drawing.Size(260, 30)
    $btnTimePeriod.Text = "TimePeriod aanmaken"
	# 3) TimePeriod aanmaken (tijdelijk console)
	$btnTimePeriod.Add_Click({
		Show-NewTimePeriodGUI -Session $Session -OwnerForm $form
	})
	
    $form.Controls.Add($btnTimePeriod)
    
    $btnAvailability = New-Object System.Windows.Forms.Button
    $btnAvailability.Location = New-Object System.Drawing.Point(310, 250)
    $btnAvailability.Size = New-Object System.Drawing.Size(260, 30)
    $btnAvailability.Text = "Availability aanmaken"
	# 4) Availability aanmaken (tijdelijk console)

	$btnAvailability.Add_Click({
		Show-NewAvailabilityGUI -Session $Session -OwnerForm $form
	})
		
    $form.Controls.Add($btnAvailability)
    
    # Sessie beheer
    $lblSession = New-Object System.Windows.Forms.Label
    $lblSession.Location = New-Object System.Drawing.Point(20, 295)
    $lblSession.Size = New-Object System.Drawing.Size(560, 25)
    $lblSession.Text = "> SESSIE BEHEER"
    $lblSession.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $lblSession.ForeColor = [System.Drawing.Color]::DarkBlue
    $form.Controls.Add($lblSession)
    
    $btnHistory = New-Object System.Windows.Forms.Button
    $btnHistory.Location = New-Object System.Drawing.Point(30, 325)
    $btnHistory.Size = New-Object System.Drawing.Size(170, 30)
    $btnHistory.Text = "History bekijken"
	# 5) History bekijken (tijdelijk console)
	$btnHistory.Add_Click({
		Show-History -Session $Session
	})
	
    $form.Controls.Add($btnHistory)
    
    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Location = New-Object System.Drawing.Point(210, 325)
    $btnSave.Size = New-Object System.Drawing.Size(170, 30)
    $btnSave.Text = "Sessie opslaan"
    $btnSave.Add_Click({
        try {
            [void](Save-Session -Session $Session)
            [System.Windows.Forms.MessageBox]::Show("Sessie succesvol opgeslagen!", "Succes", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
            
            # Update stats
            $lblActions.Text = "Aantal acties: $($Session.History.Count)"
            $tpCount = if ($Session.CreatedTimePeriods) { $Session.CreatedTimePeriods.Count } else { 0 }
            $avCount = if ($Session.CreatedAvailabilities) { $Session.CreatedAvailabilities.Count } else { 0 }
            $promCount = if ($Session.CreatedPromotions) { $Session.CreatedPromotions.Count } else { 0 }
            $lblStats.Text = "TP: $tpCount | Avail: $avCount | Prom: $promCount"
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Fout bij opslaan: $_", "Fout", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $form.Controls.Add($btnSave)
    
    $btnLoadSession = New-Object System.Windows.Forms.Button
    $btnLoadSession.Location = New-Object System.Drawing.Point(390, 325)
    $btnLoadSession.Size = New-Object System.Drawing.Size(180, 30)
    $btnLoadSession.Text = "Andere sessie laden"
	# 8) Andere sessie laden: hier WEL even menu updaten, maar zonder extra instances
	$btnLoadSession.Add_Click({
		$newSession = Show-SessionSelectionMenu
		if ($newSession) {
			$script:currentSession = $newSession
			# Sluit huidige menu en start opnieuw 1x met nieuwe session
			$form.Tag = $newSession
			$form.Close()
		}
	})

    $form.Controls.Add($btnLoadSession)
    
    $btnSessionInfo = New-Object System.Windows.Forms.Button
    $btnSessionInfo.Location = New-Object System.Drawing.Point(30, 365)
    $btnSessionInfo.Size = New-Object System.Drawing.Size(260, 30)
    $btnSessionInfo.Text = "Sessie informatie"
	# 7) Sessie info (tijdelijk console)
	$btnSessionInfo.Add_Click({
		Show-SessionInfo -Session $Session
	})
	
    $form.Controls.Add($btnSessionInfo)
    
    # Overige
    $lblOther = New-Object System.Windows.Forms.Label
    $lblOther.Location = New-Object System.Drawing.Point(20, 410)
    $lblOther.Size = New-Object System.Drawing.Size(560, 25)
    $lblOther.Text = "> OVERIGE"
    $lblOther.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $lblOther.ForeColor = [System.Drawing.Color]::DarkBlue
    $form.Controls.Add($lblOther)
    
    $btnCommitStores = New-Object System.Windows.Forms.Button
    $btnCommitStores.Location = New-Object System.Drawing.Point(30, 440)
    $btnCommitStores.Size = New-Object System.Drawing.Size(540, 30)
    $btnCommitStores.Text = "Commit Provisioning naar Stores"
    $btnCommitStores.BackColor = [System.Drawing.Color]::Orange
    $btnCommitStores.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Bold)
	# 6) Commit stores (tijdelijk console)
	$btnCommitStores.Add_Click({
		Invoke-CommitStores -Session $Session
	})
	
    $form.Controls.Add($btnCommitStores)
    
    $btnRefresh = New-Object System.Windows.Forms.Button
    $btnRefresh.Location = New-Object System.Drawing.Point(30, 475)
    $btnRefresh.Size = New-Object System.Drawing.Size(540, 30)
    $btnRefresh.Text = "Cache vernieuwen (referentiedata opnieuw ophalen)"
    $btnRefresh.ForeColor = [System.Drawing.Color]::DarkGray
    $btnRefresh.Add_Click({
        $script:cachedStores = $null
        $script:cachedSalesPoints = $null
        $script:cachedTurnoverGroups = $null
        $script:cachedVatTariffs = $null
        $script:cachedCustomers = $null
        $script:cachedProducts = $null
        
        [System.Windows.Forms.MessageBox]::Show("Cache succesvol geleegd!`nData wordt opnieuw opgehaald bij volgende gebruik.", "Cache vernieuwd", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
    $form.Controls.Add($btnRefresh)
    
    # Console/GUI toggle
    $btnToggleMode = New-Object System.Windows.Forms.Button
    $btnToggleMode.Location = New-Object System.Drawing.Point(30, 520)
    $btnToggleMode.Size = New-Object System.Drawing.Size(260, 30)
    $btnToggleMode.Text = "Schakel naar Console Modus"
    $btnToggleMode.Add_Click({
        Show-MainMenu -Session $Session
    })
    $form.Controls.Add($btnToggleMode)
    
    # Afsluiten
    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Location = New-Object System.Drawing.Point(310, 565)
    $btnExit.Size = New-Object System.Drawing.Size(260, 40)
    $btnExit.Text = "AFSLUITEN"
    $btnExit.BackColor = [System.Drawing.Color]::LightCoral
    $btnExit.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $btnExit.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show("Weet u zeker dat u wilt afsluiten?`nSessie wordt automatisch opgeslagen.", "Afsluiten", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                [void](Save-Session -Session $Session)
                Write-Log "Sessie opgeslagen voor afsluiten" -level SUCCESS
            }
            catch {
                Write-Log "FOUT bij opslaan sessie voor afsluiten: $_" -level ERROR
            }
            $form.Close()
        }
    })
    $form.Controls.Add($btnExit)
    
    # Help knop
    $btnHelp = New-Object System.Windows.Forms.Button
    $btnHelp.Location = New-Object System.Drawing.Point(30, 565)
    $btnHelp.Size = New-Object System.Drawing.Size(260, 40)
    $btnHelp.Text = "? HELP / INFO"
    $btnHelp.BackColor = [System.Drawing.Color]::LightBlue
    $btnHelp.Add_Click({
        $helpText = @"
BOOQ CM STAMDATA ONDERHOUD TOOL
================================

Versie: 1.3 (GUI Mode - Critical Bugfix Release)
Omgeving: $environment
Enterprise ID: $(Get-EnterpriseId)

FUNCTIONALITEIT:
- Promoties aanmaken en koppelen aan producten
- TimePeriods en Availabilities beheren
- Sessie-based werken (alle wijzigingen worden bewaard)
- Multi-select voor producten (Windows Forms)
- Automatische opslag en historie

TIPS:
- Gebruik 'Sessie opslaan' regelmatig
- Check 'History bekijken' voor overzicht van acties
- 'Cache vernieuwen' als data verouderd lijkt
- Schakel tussen GUI en Console modus indien gewenst

Voor ondersteuning: check de logs in de script directory
"@
        [System.Windows.Forms.MessageBox]::Show($helpText, "Help & Informatie", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
    })
    $form.Controls.Add($btnHelp)
    
    # Toon form met error handling
    try {
        $form.Add_Shown({$form.Activate()})
        [void]$form.ShowDialog()
    }
    catch {
        Write-Log "FOUT bij tonen GUI: $_" -level ERROR
        Write-Host "`n[FOUT] GUI kon niet worden getoond: $_" -ForegroundColor Red
        Write-Host "Schakel over naar Console modus..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
        Show-MainMenu -Session $Session
    }
    finally {
        if ($form) {
            $form.Dispose()
        }
    }
}



function Show-MainMenu {
    param([PSCustomObject]$Session)
    
    Clear-Host
    Write-Host "`n+==============================================================" -ForegroundColor Cyan
    Write-Host "|        BOOQ CM STAMDATA ONDERHOUD - HOOFDMENU                |" -ForegroundColor Cyan
    Write-Host "+==============================================================+" -ForegroundColor Cyan
    
    Write-Host "`nHuidige sessie: " -NoNewline -ForegroundColor Yellow
    Write-Host $Session.Name -ForegroundColor White
    Write-Host "Omgeving: " -NoNewline -ForegroundColor Yellow
    Write-Host $environment -ForegroundColor White
    Write-Host "Aantal acties: " -NoNewline -ForegroundColor Yellow
    Write-Host $Session.History.Count -ForegroundColor White
    
    Write-Host "`n+==============================================================" -ForegroundColor DarkCyan
    Write-Host "|  PROMOTIES                                                   |" -ForegroundColor DarkCyan
    Write-Host "+==============================================================+" -ForegroundColor DarkCyan
    Write-Host "1.  Promotie toevoegen aan product (VOLLEDIGE WIZARD)" -ForegroundColor Green
    Write-Host "2.  Nieuwe promotie aanmaken" -ForegroundColor White
    
    Write-Host "`n+==============================================================" -ForegroundColor DarkCyan
    Write-Host "|  ONDERSTEUNENDE ENTITEITEN                                   |" -ForegroundColor DarkCyan
    Write-Host "+==============================================================+" -ForegroundColor DarkCyan
    Write-Host "3.  TimePeriod aanmaken" -ForegroundColor White
    Write-Host "4.  Availability aanmaken" -ForegroundColor White
    
    Write-Host "`n+==============================================================" -ForegroundColor DarkCyan
    Write-Host "|  SESSIE BEHEER                                               |" -ForegroundColor DarkCyan
    Write-Host "+==============================================================+" -ForegroundColor DarkCyan
    Write-Host "5.  Sessie history bekijken" -ForegroundColor Yellow
    Write-Host "6.  Sessie opslaan" -ForegroundColor Yellow
    Write-Host "7.  Andere sessie laden" -ForegroundColor Yellow
    
    Write-Host "`n+==============================================================" -ForegroundColor DarkCyan
    Write-Host "|  OVERIGE                                                     |" -ForegroundColor DarkCyan
    Write-Host "+==============================================================+" -ForegroundColor DarkCyan
    Write-Host "8.  Commit Provisioning naar Stores" -ForegroundColor Magenta
    Write-Host "9.  Cache vernieuwen (opnieuw ophalen referentiedata)" -ForegroundColor Gray
    Write-Host "10. Sessie informatie tonen" -ForegroundColor Gray
    Write-Host "11. Schakel naar GUI Modus" -ForegroundColor Cyan
    
    Write-Host "`nQ.  Afsluiten`n" -ForegroundColor Red
    
    $choice = Read-Host "Maak uw keuze"
    
    switch ($choice.ToUpper()) {
        "1" {
            Add-PromotionToProduct -Session $Session
            Show-MainMenu -Session $Session
        }
        "2" {
            New-Promotion -Session $Session
            Show-MainMenu -Session $Session
        }
        "3" {
            New-TimePeriod -Session $Session
            Show-MainMenu -Session $Session
        }
        "4" {
            New-Availability -Session $Session
            Show-MainMenu -Session $Session
        }
        "5" {
            Show-History -Session $Session
            Show-MainMenu -Session $Session
        }
        "6" {
            try {
                [void](Save-Session -Session $Session)
                Write-Host "`n[OK] Sessie opgeslagen!" -ForegroundColor Green
            }
            catch {
                Write-Host "`n[FOUT] Sessie kon niet worden opgeslagen: $_" -ForegroundColor Red
            }
            Start-Sleep -Seconds 1
            Show-MainMenu -Session $Session
        }
        "7" {
            $newSession = Show-SessionSelectionMenu
            if ($newSession) {
                $script:currentSession = $newSession
                Show-MainMenu -Session $newSession
            }
            else {
                Show-MainMenu -Session $Session
            }
        }
        "8" {
            Invoke-CommitStores -Session $Session
            Show-MainMenu -Session $Session
        }
        "9" {
            Clear-Host
            Write-Host "`n=== CACHE VERNIEUWEN ===" -ForegroundColor Cyan
            Write-Host "`nCache wordt geleegd..." -ForegroundColor Yellow
            
            $script:cachedStores = $null
            $script:cachedSalesPoints = $null
            $script:cachedTurnoverGroups = $null
            $script:cachedVatTariffs = $null
            $script:cachedCustomers = $null
            $script:cachedProducts = $null
            
            Write-Host "[OK] Cache succesvol geleegd!" -ForegroundColor Green
            Write-Host "Data wordt opnieuw opgehaald bij volgende gebruik." -ForegroundColor Gray
            
            Read-Host "`nDruk op Enter om door te gaan"
            Show-MainMenu -Session $Session
        }
        "10" {
            Show-SessionInfo -Session $Session
            Show-MainMenu -Session $Session
        }
        "11" {
        }
        "Q" {
            Write-Host "`nSessie wordt opgeslagen..." -ForegroundColor Yellow
            try {
                [void](Save-Session -Session $Session)
                Write-Host "[OK] Sessie opgeslagen!" -ForegroundColor Green
            }
            catch {
                Write-Host "[WAARSCHUWING] Sessie kon niet worden opgeslagen: $_" -ForegroundColor Yellow
            }
            Write-Host "Tot ziens!" -ForegroundColor Green
            Write-Log "Script afgesloten door gebruiker" -Level INFO
            exit 0
        }
        default {
            Write-Host "`n[FOUT] Ongeldige keuze. Probeer opnieuw." -ForegroundColor Red
            Start-Sleep -Seconds 1
            Show-MainMenu -Session $Session
        }
    }
}

function Show-SessionInfo {
    param([PSCustomObject]$Session)
    
    Clear-Host
    Write-Host "`n+==============================================================" -ForegroundColor Cyan
    Write-Host "|                 SESSIE INFORMATIE                            |" -ForegroundColor Cyan
    Write-Host "+==============================================================+" -ForegroundColor Cyan
    
    Write-Host "`nSessie Details:" -ForegroundColor Yellow
    Write-Host "  ID: $($Session.SessionId)" -ForegroundColor Gray
    Write-Host "  Naam: $($Session.Name)" -ForegroundColor White
    Write-Host "  Beschrijving: $($Session.Description)" -ForegroundColor White
    Write-Host "  Aangemaakt: $($Session.CreatedAt)" -ForegroundColor Gray
    Write-Host "  Laatst gewijzigd: $($Session.LastModified)" -ForegroundColor Gray
    Write-Host "  Omgeving: $($Session.Environment)" -ForegroundColor White
    
    Write-Host "`nGemaakte Entiteiten:" -ForegroundColor Yellow
    
    # TimePeriods
    $tpCount = if ($Session.CreatedTimePeriods) { $Session.CreatedTimePeriods.Count } else { 0 }
    Write-Host "  TimePeriods: $tpCount" -ForegroundColor White
    if ($tpCount -gt 0) {
        foreach ($tp in $Session.CreatedTimePeriods) {
            Write-Host "    - $tp" -ForegroundColor DarkGray
        }
    }
    
    # Availabilities
    $avCount = if ($Session.CreatedAvailabilities) { $Session.CreatedAvailabilities.Count } else { 0 }
    Write-Host "  Availabilities: $avCount" -ForegroundColor White
    if ($avCount -gt 0) {
        foreach ($av in $Session.CreatedAvailabilities) {
            Write-Host "    - $av" -ForegroundColor DarkGray
        }
    }
    
    # Customers
    $custCount = if ($Session.CreatedCustomers) { $Session.CreatedCustomers.Count } else { 0 }
    Write-Host "  Customers: $custCount" -ForegroundColor White
    if ($custCount -gt 0) {
        foreach ($cust in $Session.CreatedCustomers) {
            Write-Host "    - $cust" -ForegroundColor DarkGray
        }
    }
    
    # Promotions
    $promCount = if ($Session.CreatedPromotions) { $Session.CreatedPromotions.Count } else { 0 }
    Write-Host "  Promotions: $promCount" -ForegroundColor White
    if ($promCount -gt 0) {
        foreach ($prom in $Session.CreatedPromotions) {
            Write-Host "    - $prom" -ForegroundColor DarkGray
        }
    }
    
    # Products
    $prodCount = if ($Session.CreatedProducts) { $Session.CreatedProducts.Count } else { 0 }
    Write-Host "  Products (geselecteerd): $prodCount" -ForegroundColor White
    if ($prodCount -gt 0) {
        foreach ($prod in $Session.CreatedProducts) {
            Write-Host "    - $prod" -ForegroundColor DarkGray
        }
    }
    
    Write-Host "`nHistory:" -ForegroundColor Yellow
    $histCount = if ($Session.History) { $Session.History.Count } else { 0 }
    Write-Host "  Totaal aantal acties: $histCount" -ForegroundColor White
    
    if ($histCount -gt 0) {
        Write-Host "`n  Laatste 5 acties:" -ForegroundColor Gray
        $lastActions = $Session.History | Select-Object -Last 5
        foreach ($action in $lastActions) {
            $timestamp = if ($action.Timestamp) { $action.Timestamp.ToString('HH:mm:ss') } else { "??:??:??" }
            Write-Host "    [$timestamp] $($action.Action) $($action.EntityType)" -ForegroundColor DarkGray
        }
    }
    
    Read-Host "`nDruk op Enter om door te gaan"
}

function Add-PromotionToProduct {
    param([PSCustomObject]$Session)
    
    Clear-Host
    Write-Host "`n+==============================================================" -ForegroundColor Cyan
    Write-Host "|         PROMOTIE TOEVOEGEN AAN PRODUCT                       |" -ForegroundColor Cyan
    Write-Host "+==============================================================+" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Deze wizard helpt u om een promotie aan een product toe te voegen." -ForegroundColor Yellow
    Write-Host "U moet de volgende stappen doorlopen:" -ForegroundColor Yellow
    Write-Host "  1. Product selecteren" -ForegroundColor Gray
    Write-Host "  2. Promotie aanmaken of selecteren" -ForegroundColor Gray
    Write-Host "  3. TimePeriod(s) aanmaken indien nodig" -ForegroundColor Gray
    Write-Host "  4. Availability(s) aanmaken indien nodig" -ForegroundColor Gray
    Write-Host "  5. Customer selectie (optioneel)" -ForegroundColor Gray
    Write-Host ""
    
    Read-Host "Druk op Enter om te starten"
    
    # Start met het aanmaken van de promotie
    # Dit proces zal alle benodigde stappen doorlopen
    $promotionId = New-Promotion -Session $Session
    
    if ($promotionId) {
        Write-Host "`n[OK] Promotie succesvol aangemaakt en gekoppeld!" -ForegroundColor Green
        Write-Host "Promotie ID: $promotionId" -ForegroundColor Cyan
    }
    else {
        Write-Host "`n[FOUT] Promotie niet aangemaakt." -ForegroundColor Red
    }
    
    Read-Host "`nDruk op Enter om terug te gaan naar het hoofdmenu"
}


function Show-AddPromotionWizardGUI {
    param(
        [Parameter(Mandatory)]
        [PSCustomObject]$Session,

        [Parameter(Mandatory=$false)]
        [System.Windows.Forms.Form]$OwnerForm
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $dlg = New-Object System.Windows.Forms.Form
    $dlg.Text = "Promotie toevoegen aan product (Wizard)"
    $dlg.StartPosition = if ($OwnerForm) { "CenterParent" } else { "CenterScreen" }
    $dlg.Size = New-Object System.Drawing.Size(760, 520)
    $dlg.FormBorderStyle = "FixedDialog"
    $dlg.MaximizeBox = $false
    $dlg.MinimizeBox = $false
    $dlg.ShowInTaskbar = $false

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Location = New-Object System.Drawing.Point(16, 16)
    $lbl.Size = New-Object System.Drawing.Size(720, 40)
    $lbl.Text = "Deze wizard wordt stap voor stap omgezet van console naar GUI. Selecteer alvast producten (PIM)."
    $dlg.Controls.Add($lbl)

    $btnPickProducts = New-Object System.Windows.Forms.Button
    $btnPickProducts.Location = New-Object System.Drawing.Point(16, 72)
    $btnPickProducts.Size = New-Object System.Drawing.Size(220, 32)
    $btnPickProducts.Text = "Producten selecteren..."
    $dlg.Controls.Add($btnPickProducts)

    $txtSelected = New-Object System.Windows.Forms.TextBox
    $txtSelected.Location = New-Object System.Drawing.Point(16, 120)
    $txtSelected.Size = New-Object System.Drawing.Size(720, 300)
    $txtSelected.Multiline = $true
    $txtSelected.ScrollBars = "Vertical"
    $txtSelected.ReadOnly = $true
    $dlg.Controls.Add($txtSelected)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Location = New-Object System.Drawing.Point(640, 440)
    $btnClose.Size = New-Object System.Drawing.Size(96, 32)
    $btnClose.Text = "Sluiten"
    $btnClose.Add_Click({ $dlg.Close() })
    $dlg.Controls.Add($btnClose)

    $selectedProductIds = @()

    $btnPickProducts.Add_Click({
        try {
            # Laad producten uit PIM (cached)
            $products = Get-CachedProducts
            if (-not $products -or $products.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show(
                    "Geen producten gevonden in PIM.",
                    "Geen data",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Warning
                ) | Out-Null
                return
            }

            $preSelected = if ($selectedProductIds) { $selectedProductIds } else { @() }
            $picked = Show-ProductMultiSelect -Products $products -PreSelectedIds $preSelected

            if ($picked -and $picked.Count -gt 0) {
                $selectedProductIds = @($picked)

                $txtSelected.Clear()
                $txtSelected.AppendText("Geselecteerde producten: $($selectedProductIds.Count)`r`n`r`n") | Out-Null
                foreach ($id in $selectedProductIds) {
                    $txtSelected.AppendText("- $id`r`n") | Out-Null
                }
            }
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Fout bij productselectie: $($_.Exception.Message)",
                "Fout",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    })

    if ($OwnerForm) {
        $null = $dlg.ShowDialog($OwnerForm)
    } else {
        $null = $dlg.ShowDialog()
    }
}


# ============================================================================
# COMMIT STORES FUNCTIE
# ============================================================================

function Invoke-CommitStores {
    <#
    .SYNOPSIS
        Commit provisioning packages naar alle stores
    #>
    param([PSCustomObject]$Session)
    
    Clear-Host
    Write-Host "`n+==============================================================" -ForegroundColor Cyan
    Write-Host "|         COMMIT PROVISIONING NAAR STORES                      |" -ForegroundColor Cyan
    Write-Host "+==============================================================+" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Deze functie commit alle provisioning wijzigingen naar de stores." -ForegroundColor Yellow
    Write-Host "Dit kan enige tijd duren afhankelijk van het aantal stores." -ForegroundColor Yellow
    Write-Host ""
    
    # Haal stores op
    Write-Host "Stores ophalen..." -ForegroundColor Yellow
    $stores = Get-CachedStores
    
    if ($stores.Count -eq 0) {
        Write-Host "`n[FOUT] Geen stores gevonden!" -ForegroundColor Red
        Read-Host "Druk op Enter om terug te gaan"
        return
    }
    
    Write-Host "[OK] $($stores.Count) store(s) gevonden" -ForegroundColor Green
    Write-Host ""
    
    # Toon stores
    Write-Host "Stores:" -ForegroundColor Yellow
    foreach ($store in $stores) {
        Write-Host "  - $($store.name) (ID: $($store.Id))" -ForegroundColor Gray
    }
    Write-Host ""
    
    # Bevestiging
    Write-Host "WAARSCHUWING: Deze actie commit wijzigingen naar ALLE stores!" -ForegroundColor Red
    $confirm = Read-Host "Weet u zeker dat u wilt doorgaan? (typ 'JA' om te bevestigen)"
    
    if ($confirm -ne 'JA') {
        Write-Host "`n[INFO] Geannuleerd" -ForegroundColor Yellow
        Read-Host "Druk op Enter om terug te gaan"
        return
    }
    
    Write-Host "`n"
    Write-Host "Commit wordt uitgevoerd..." -ForegroundColor Yellow
    Write-Host ""
    
    $successCount = 0
    $failCount = 0
    $results = @()
    
    foreach ($store in $stores) {
        $storeId = $store.Id
        $storeName = $store.name
        
        Write-Host "Commit voor store: $storeName..." -ForegroundColor Cyan
        
        try {
            $response = Invoke-CMApi -Endpoint "stores/$storeId/commit" -Method PUT
            Write-Host "  [OK] $storeName - Commit succesvol" -ForegroundColor Green
            $successCount++
            
            $results += [PSCustomObject]@{
                Store = $storeName
                StoreId = $storeId
                Status = "Success"
                Message = "Commit succesvol"
            }
            
            # Log in sessie history
            $Session.History += @{
                Timestamp = Get-Date
                Action = "Commit"
                EntityType = "Store"
                EntityId = $storeId
                EntityName = $storeName
            }
        }
        catch {
            Write-Host "  [FOUT] $storeName - Commit gefaald: $_" -ForegroundColor Red
            Write-Log "Commit gefaald voor store $storeName ($storeId): $_" -level ERROR
            $failCount++
            
            $results += [PSCustomObject]@{
                Store = $storeName
                StoreId = $storeId
                Status = "Failed"
                Message = $_.Exception.Message
            }
        }
        
        Start-Sleep -Milliseconds 500
    }
    
    # Samenvatting
    Write-Host "`n"
    Write-Host "+==============================================================" -ForegroundColor Cyan
    Write-Host "|                    COMMIT SAMENVATTING                       |" -ForegroundColor Cyan
    Write-Host "+==============================================================+" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Totaal stores: $($stores.Count)" -ForegroundColor White
    Write-Host "Succesvol: $successCount" -ForegroundColor Green
    Write-Host "Gefaald: $failCount" -ForegroundColor Red
    Write-Host ""
    
    if ($failCount -gt 0) {
        Write-Host "Gefaalde stores:" -ForegroundColor Red
        foreach ($result in $results | Where-Object { $_.Status -eq "Failed" }) {
            Write-Host "  - $($result.Store): $($result.Message)" -ForegroundColor Gray
        }
        Write-Host ""
    }
    
    # Sessie opslaan
    Save-Session -Session $Session | Out-Null
    Write-Log "Commit stores uitgevoerd: $successCount succesvol, $failCount gefaald" -level INFO
    
    Read-Host "`nDruk op Enter om terug te gaan"
}

# ============================================================================
# MAIN
# ============================================================================
Write-Host "`n+==============================================================" -ForegroundColor Green
Write-Host "|                                                               |" -ForegroundColor Green
Write-Host "|      WELKOM BIJ BOOQ CM STAMDATA ONDERHOUD TOOL              |" -ForegroundColor Green
Write-Host "|                                                               |" -ForegroundColor Green
Write-Host "+==============================================================+" -ForegroundColor Green
Write-Host ""
Write-Host "Versie: 1.2 (Bugfixes: Forms, Encoding, JSON export, Commit Stores)" -ForegroundColor Gray
Write-Host "Omgeving: $environment" -ForegroundColor Yellow
if ($showExtraDetails) {
    Write-Host "Extra details: INGESCHAKELD" -ForegroundColor Magenta
    Write-Host "API requests worden gelogd naar: CM-API-Requests.csv" -ForegroundColor Magenta
}
Write-Host ""

# Test authenticatie
Write-Host "Authenticatie wordt getest..." -ForegroundColor Yellow
try {
    $token = Get-ValidToken -ClientId $clientId -ClientSecret $clientSecret -TokenUrl $script:tokenUrl -ImpersonateClientId $impersonateClientId -DebugMode $DebugMode
    $payload = Get-JWTPayload -Token $token
    $enterpriseId = $payload.enterprise_id
    
    Write-Host "[OK] Authenticatie succesvol!" -ForegroundColor Green
    Write-Host "  Enterprise ID: $enterpriseId" -ForegroundColor Cyan
    
    if ($impersonateClientId) {
        Write-Host "  Impersoneren als: $impersonateClientId" -ForegroundColor Cyan
    }
    
    Write-Log "Authenticatie succesvol - Enterprise ID: $enterpriseId" -Level SUCCESS
}
catch {
    Write-Host "[FOUT] Authenticatie gefaald!" -ForegroundColor Red
    Write-Host "  Fout: $_" -ForegroundColor Red
    Write-Log "Authenticatie gefaald: $_" -Level ERROR
    exit 1
}

# Initialiseer API URLs
Write-Host "`nAPI URLs initialiseren..." -ForegroundColor Yellow
if (-not (Initialize-ApiUrls)) {
    Write-Host "[FOUT] Kan API URLs niet initialiseren!" -ForegroundColor Red
    exit 1
}

# Als onboarding mode, haal data op en toon
if ($onboarding) {
    Write-Host "`n=== ONBOARDING MODUS ===" -ForegroundColor Cyan
    
    Write-Host "`nStores ophalen..." -ForegroundColor Yellow
    $stores = Get-CachedStores
    Write-Host "[OK] Stores: $($stores.Count)" -ForegroundColor Green
    
    Write-Host "`nSalesPoints ophalen..." -ForegroundColor Yellow
    $salesPoints = Get-CachedSalesPoints
    Write-Host "[OK] SalesPoints: $($salesPoints.Count)" -ForegroundColor Green
    
    Write-Host "`nTurnoverGroups ophalen..." -ForegroundColor Yellow
    $tg = Get-CachedTurnoverGroups
    Write-Host "[OK] TurnoverGroups: $($tg.Count)" -ForegroundColor Green
    
    Write-Host "`nVatTariffs ophalen..." -ForegroundColor Yellow
    $vt = Get-CachedVatTariffs
    Write-Host "[OK] VatTariffs: $($vt.Count)" -ForegroundColor Green
    
    Write-Host "`nCustomers ophalen..." -ForegroundColor Yellow
    $c = Get-CachedCustomers
    Write-Host "[OK] Customers: $($c.Count)" -ForegroundColor Green
    
    Write-Host "`nProducts ophalen..." -ForegroundColor Yellow
    $p = Get-CachedProducts
    Write-Host "[OK] Products: $($p.Count)" -ForegroundColor Green
    
    Write-Host "`n=== ONBOARDING DATA SUCCESVOL OPGEHAALD ===" -ForegroundColor Green
    Read-Host "Druk op Enter om verder te gaan naar het hoofdmenu"
}

Write-Host ""
Read-Host "Druk op Enter om door te gaan"

# Selecteer of maak sessie
try {
    $script:currentSession = Show-SessionSelectionMenu
    
    if ($script:currentSession) {
        Write-Log "Sessie geselecteerd: $($script:currentSession.Name)" -Level INFO
        Show-MainMenuGUI -Session $script:currentSession
    }
    else {
        Write-Host "`n[INFO] Geen sessie geselecteerd. Script wordt afgesloten." -ForegroundColor Yellow
        Write-Log "Script afgesloten - geen sessie geselecteerd" -Level INFO
        exit 0
    }
}
catch {
    Write-Host "`n[FOUT] Er is een fout opgetreden: $_" -ForegroundColor Red
    Write-Host "Details: $($_.Exception.Message)" -ForegroundColor Red
    Write-Log "Onverwachte fout in main: $_" -Level ERROR
    
    Read-Host "`nDruk op Enter om af te sluiten"
    exit 1
}
finally {
    Set-Location $script:startLocation
    Write-Log "=== Script afgesloten ===" -level INFO
    Write-Host "`nStart locatie hersteld: $($script:startLocation)" -ForegroundColor Cyan
    if ($showExtraDetails) {
        $logFilePath = Join-Path $script:startLocation "CM-API-Requests.csv"
        if (Test-Path $logFilePath) {
            Write-Host "API request log: $logFilePath" -ForegroundColor Magenta
        }
    }
}





