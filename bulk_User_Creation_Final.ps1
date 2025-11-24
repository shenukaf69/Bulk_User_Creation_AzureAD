<# ==================================================================== 
 Script Name : bulk_User_Creation_Final.ps1
 Purpose     : Bulk create users (skip if existing), assign E1/E3 + Teams licenses,
               enable archive mailbox (if not already enabled, after mailbox provisioned),
               then enable auto-expanding archive,
               and export reports.
 Tenant      : <your-tenant-name>.onmicrosoft.com
 Requirements:
   • PowerShell 7+
   • Microsoft.Graph (v2+)
   • ExchangeOnlineManagement
 CSV Format  : bulk_User_creation_file.csv must include:
               Source Users UPN, Target_UPN, Display name, Job title, Department,
               Password, License, Archive
 ==================================================================== #>

# -------------------------------------------------------------
# Paths (use repo-relative paths so nothing tenant-specific)
# -------------------------------------------------------------
$csvPath        = ".\bulk_User_creation_file.csv"
$exportMain     = ".\Output\User_License_Report_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
$exportSkipped  = ".\Output\User_Skipped_{0}.csv" -f (Get-Date -Format "yyyyMMdd_HHmmss")
$logFile        = ".\Output\UserCreationLog_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")

# Ensure Output folder exists
New-Item -ItemType Directory -Path (Split-Path $exportMain) -Force | Out-Null

# -------------------------------------------------------------
# Import users from CSV, skip invalid rows
# -------------------------------------------------------------
$users = Import-Csv -Path $csvPath | Where-Object {
    $_.Target_UPN -and $_.'Display name' -and $_.License -and
    $_.Target_UPN -ne '#N/A' -and $_.License -ne '#N/A' -and $_.'Display name' -ne '#N/A'
}

# -------------------------------------------------------------
# Logging helper
# -------------------------------------------------------------
New-Item -Path $logFile -ItemType File -Force | Out-Null
function Write-Log {
    param([string]$message)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$timestamp - $message" | Out-File -FilePath $logFile -Append
    Write-Host $message
}

# -------------------------------------------------------------
# Connect to Microsoft Graph (Device Code auth)
# -------------------------------------------------------------
try {
    Write-Log "🔄 Connecting to Microsoft Graph using Device Code..."
    Connect-MgGraph -Scopes "User.ReadWrite.All","Directory.ReadWrite.All" -UseDeviceCode -NoWelcome
    Write-Log "✅ Connected to Microsoft Graph."
}
catch {
    Write-Log "❌ Failed to connect to Microsoft Graph: $($_.Exception.Message)"
    exit
}

# -------------------------------------------------------------
# Connect to Exchange Online
# -------------------------------------------------------------
try {
    Write-Log "🔄 Connecting to Exchange Online..."
    # Use interactive auth or your own admin UPN; do NOT hardcode real tenant UPNs in public repos
    Connect-ExchangeOnline                            # or: Connect-ExchangeOnline -UserPrincipalName "<admin-upn>@<your-tenant>.onmicrosoft.com"
    Write-Log "✅ Connected to Exchange Online."
}
catch {
    Write-Log "❌ Failed to connect to Exchange Online: $($_.Exception.Message)"
    exit
}

# -------------------------------------------------------------
# Detect license SKUs automatically
# -------------------------------------------------------------
Write-Log "🔍 Checking available license counts..."
$licensedSkus = Get-MgSubscribedSku | Select SkuId, SkuPartNumber, ConsumedUnits, `
    @{n='AvailableUnits';e={$_.PrepaidUnits.Enabled - $_.ConsumedUnits}}

function Get-SkuId([string]$pattern) {
    $m = $licensedSkus | Where-Object { $_.SkuPartNumber -match $pattern } | Select -First 1
    if (-not $m) { throw "❌ Can't find SKU matching pattern: $pattern. Available: $($licensedSkus.SkuPartNumber -join ', ')" }
    return $m.SkuId
}

# Auto-detect SKUs
$SkuE3    = Get-SkuId '(^|_)E3($|_)|ENTERPRISEPACK|OFFICE.*E3'
$SkuE1    = Get-SkuId '(^|_)E1($|_)|STANDARDPACK|OFFICE.*E1'
$SkuTeams = Get-SkuId 'TEAMS.*ENTERPRISE|MICROSOFT_TEAMS(_ENTERPRISE)?'

$availableE3    = ($licensedSkus | Where-Object { $_.SkuId -eq $SkuE3 }).AvailableUnits
$availableE1    = ($licensedSkus | Where-Object { $_.SkuId -eq $SkuE1 }).AvailableUnits
$availableTeams = ($licensedSkus | Where-Object { $_.SkuId -eq $SkuTeams }).AvailableUnits

Write-Log "📦 Resolved SKUs -> E1: $SkuE1 | E3: $SkuE3 | Teams: $SkuTeams"
Write-Log "📦 Available -> E1: $availableE1 | E3: $availableE3 | Teams: $availableTeams"

# -------------------------------------------------------------
# Variables and arrays
# -------------------------------------------------------------
$report       = @()
$skippedUsers = @()

# -------------------------------------------------------------
# Process each user
# -------------------------------------------------------------
foreach ($user in $users) {
    $upn         = $user.Target_UPN.Trim()
    $displayName = $user.'Display name'
    $licenseType = $user.License.Trim().ToUpper()
    $password    = $user.Password
    $jobTitle    = $user.'Job title'
    $department  = $user.Department
    $archivePref = $user.Archive
    $status      = "Pending"
    $archive     = "Not Enabled"

    try {
        Write-Log "➡️ Processing $upn ($licenseType)"

        # -----------------------------------------------------
        # Check if user already exists
        # -----------------------------------------------------
        $existingUser = Get-MgUser -UserId $upn -ErrorAction SilentlyContinue
        if ($existingUser) {
            Write-Log "ℹ️ User already exists: $upn — skipping creation."
        }
        else {
            $userParams = @{
                DisplayName       = $displayName
                UserPrincipalName = $upn
                MailNickname      = ($upn.Split("@")[0])
                PasswordProfile   = @{ ForceChangePasswordNextSignIn = $true; Password = $password }
                Department        = $department
                JobTitle          = $jobTitle
                AccountEnabled    = $true
            }
            New-MgUser @userParams
            Write-Log "✅ User created: $upn"
        }

        # -----------------------------------------------------
        # Set usage location
        # -----------------------------------------------------
        Update-MgUser -UserId $upn -UsageLocation "SG"

        # -----------------------------------------------------
        # Assign license
        # -----------------------------------------------------
        $licensesToAssign = @()
        switch ($licenseType) {
            "E3" {
                if ($availableE3 -gt 0 -and $availableTeams -gt 0) {
                    $licensesToAssign += @{ SkuId = $SkuE3 }
                    $licensesToAssign += @{ SkuId = $SkuTeams }
                    $availableE3--; $availableTeams--
                    Set-MgUserLicense -UserId $upn -AddLicenses $licensesToAssign -RemoveLicenses @()
                    $status = "E3 + Teams assigned"
                    Write-Log "✅ $status for $upn"
                }
                else {
                    $status = "❌ Skipped (E3/Teams license exhausted)"
                    $skippedUsers += [PSCustomObject]@{Target_UPN=$upn;Reason="No E3/Teams license left"}
                    Write-Log $status
                }
            }
            "E1" {
                if ($availableE1 -gt 0 -and $availableTeams -gt 0) {
                    $licensesToAssign += @{ SkuId = $SkuE1 }
                    $licensesToAssign += @{ SkuId = $SkuTeams }
                    $availableE1--; $availableTeams--
                    Set-MgUserLicense -UserId $upn -AddLicenses $licensesToAssign -RemoveLicenses @()
                    $status = "E1 + Teams assigned"
                    Write-Log "✅ $status for $upn"
                }
                else {
                    $status = "❌ Skipped (E1/Teams license exhausted)"
                    $skippedUsers += [PSCustomObject]@{Target_UPN=$upn;Reason="No E1/Teams license left"}
                    Write-Log $status
                }
            }
            default {
                $status = "⚠️ Unknown license type ($licenseType)"
                Write-Log $status
            }
        }

        # -----------------------------------------------------
        # Enable archive mailbox if required
        # -----------------------------------------------------
        if ($archivePref -eq "Yes" -and $status -like "*assigned*") {
            # Check if mailbox already exists
            $existingMailbox = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
            if ($existingMailbox) {
                Write-Log "ℹ️ Mailbox already exists for $upn — skipping archive creation."
            }
            else {
                Write-Log "⏳ Waiting for mailbox to provision (up to 8 minutes)..."
                $timeout = (Get-Date).AddMinutes(8)
                $mailboxReady = $false

                do {
                    Start-Sleep -Seconds 30
                    $mbx = Get-Mailbox -Identity $upn -ErrorAction SilentlyContinue
                    if ($mbx) { $mailboxReady = $true }
                } until ($mailboxReady -or (Get-Date) -gt $timeout)

                if ($mailboxReady) {
                    Enable-Mailbox -Identity $upn -Archive -ErrorAction Stop
                    Write-Log "📦 Archive mailbox enabled for $upn"

                    Write-Log "⏳ Waiting 2 minutes before enabling Auto-Expanding Archive..."
                    Start-Sleep -Seconds 120

                    Enable-Mailbox -Identity $upn -AutoExpandingArchive -ErrorAction SilentlyContinue
                    $archive = "Enabled (Auto-expanding)"
                    Write-Log "📈 Auto-expanding archive enabled for $upn"
                }
                else {
                    $archive = "⚠️ Mailbox not found after 8 minutes — skipping archive"
                    Write-Log $archive
                }
            }
        }

        # -----------------------------------------------------
        # Append results
        # -----------------------------------------------------
        $report += [PSCustomObject]@{
            Target_UPN   = $upn
            DisplayName  = $displayName
            LicenseType  = $licenseType
            Status       = $status
            Archive      = $archive
        }
    }
    catch {
        Write-Log "❌ Error on $upn — $($_.Exception.Message)"
        $report += [PSCustomObject]@{
            Target_UPN   = $upn
            DisplayName  = $displayName
            LicenseType  = $licenseType
            Status       = "Failed"
            Archive      = "N/A"
        }
    }
}

# -------------------------------------------------------------
# Export results
# -------------------------------------------------------------
$report | Export-Csv -Path $exportMain -NoTypeInformation -Encoding UTF8
if ($skippedUsers.Count -gt 0) {
    $skippedUsers | Export-Csv -Path $exportSkipped -NoTypeInformation -Encoding UTF8
    Write-Log "⚠️ Skipped users exported to $exportSkipped"
}

Write-Log "`n📊 Main report exported to $exportMain"
Write-Log "📦 Remaining -> E3: $availableE3 | E1: $availableE1 | Teams: $availableTeams"
Write-Log "🎯 All users processed. Log: $logFile"
