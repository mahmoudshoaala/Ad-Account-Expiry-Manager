#Requires -Modules ActiveDirectory
<#
.SYNOPSIS
    AD Account Expiry Manager - Business Exception Handler
.DESCRIPTION
    GUI tool for managing business exception accounts terminated in HR
    but still requiring O365 access for a defined period.
    Supports dual DC failover: Contoso-dcsrv-01 / Contoso-dcsrv-02
.NOTES
    Version : 4.0 (SMTP SmtpClient fix - 2026-03-08)
    Requires: PowerShell 5.1, ActiveDirectory module, RSAT
#>

# ---------------------------------------------------------------------------
# CONFIGURATION -- Edit before deployment
# ---------------------------------------------------------------------------
$Global:Config = @{
    PrimaryDC        = "Contoso-dcsrv-01"
    SecondaryDC      = "Contoso-dcsrv-02"
    DisabledOU       = "OU=Disabled Accounts,DC=yourdomain,DC=com"
    # ---- SMTP Settings ----
    SMTPServer       = "smtp.yourdomain.com"   # e.g. mail.contoso.com or 192.168.1.10
    SMTPPort         = 25                       # 25 = relay, 587 = TLS submission, 465 = SSL
    SMTPUseSSL       = $false                  # Set $true for port 587/465
    SMTPFrom         = "ad-automation@yourdomain.com"
    SMTPUser         = ""                       # Leave empty for anonymous relay
    SMTPPass         = ""                       # Leave empty for anonymous relay
    ReportRecipients = @("it-admin@yourdomain.com", "security@yourdomain.com")
    # -----------------------
    DataFile         = "$PSScriptRoot\ExceptionAccounts.xml"
    LogFile          = "$PSScriptRoot\Logs\ExpiryManager_$(Get-Date -f 'yyyy-MM').log"
    ReportDir        = "$PSScriptRoot\Reports"
}

# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $logDir = Split-Path $Global:Config.LogFile
    if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
    $entry = "[$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')] [$Level] $Message"
    Add-Content -Path $Global:Config.LogFile -Value $entry -Encoding UTF8
    if ($Level -eq "ERROR") { Write-Host $entry -ForegroundColor Red }
    if ($Level -eq "WARN")  { Write-Host $entry -ForegroundColor Yellow }
}

# ---------------------------------------------------------------------------
# DATA PERSISTENCE
# ---------------------------------------------------------------------------
function Load-AccountList {
    if (Test-Path $Global:Config.DataFile) {
        try {
            $raw = Import-Clixml -Path $Global:Config.DataFile
            return @($raw)
        } catch {
            Write-Log "Failed to load data file: $_" "ERROR"
            return @()
        }
    }
    return @()
}

function Save-AccountList {
    param([array]$List)
    try {
        # Ensure directory exists
        $dir = Split-Path $Global:Config.DataFile
        if ($dir -and -not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $List | Export-Clixml -Path $Global:Config.DataFile -Force
        Write-Log "Account list saved. Total entries: $($List.Count)"
    } catch {
        Write-Log "Failed to save data file: $_" "ERROR"
    }
}

# ---------------------------------------------------------------------------
# ACTIVE DIRECTORY HELPERS
# ---------------------------------------------------------------------------
function Get-AvailableDC {
    foreach ($dc in @($Global:Config.PrimaryDC, $Global:Config.SecondaryDC)) {
        if (Test-Connection -ComputerName $dc -Count 1 -Quiet -ErrorAction SilentlyContinue) {
            Write-Log "Using domain controller: $dc"
            return $dc
        }
    }
    throw "Neither domain controller is reachable ($($Global:Config.PrimaryDC), $($Global:Config.SecondaryDC))"
}

function Get-ADUserSafe {
    param([string]$SamAccountName, [string]$DC)
    try {
        return Get-ADUser -Identity $SamAccountName `
                          -Server $DC `
                          -Properties DisplayName,Enabled,DistinguishedName,
                                      AccountExpirationDate,EmailAddress,Department,Title `
                          -ErrorAction Stop
    } catch { return $null }
}

function Disable-ExceptionAccount {
    param([string]$SamAccountName, [datetime]$ExpirationDate, [string]$DC)

    $result = [PSCustomObject]@{
        SamAccountName = $SamAccountName
        DisplayName    = ""
        ExpirationDate = $ExpirationDate
        Action         = ""
        Status         = ""
        ErrorMessage   = ""
        Timestamp      = Get-Date
        DCUsed         = $DC
    }

    try {
        $user = Get-ADUserSafe -SamAccountName $SamAccountName -DC $DC
        if (-not $user) {
            $result.Status       = "FAILED"
            $result.ErrorMessage = "User not found in Active Directory"
            Write-Log "User not found: $SamAccountName" "ERROR"
            return $result
        }

        $result.DisplayName = [string]$user.DisplayName
        $prevOU = ($user.DistinguishedName -replace '^CN=[^,]+,', '')
        $actions = @()

        Set-ADAccountExpiration -Identity $SamAccountName -Server $DC -DateTime $ExpirationDate -ErrorAction Stop
        $actions += "Expiration set to $($ExpirationDate.ToString('yyyy-MM-dd'))"
        Write-Log "Expiration set for $SamAccountName -> $($ExpirationDate.ToString('yyyy-MM-dd'))"

        Disable-ADAccount -Identity $SamAccountName -Server $DC -ErrorAction Stop
        $actions += "Account disabled"
        Write-Log "Account disabled: $SamAccountName"

        if ($prevOU -ne $Global:Config.DisabledOU) {
            Move-ADObject -Identity $user.DistinguishedName -TargetPath $Global:Config.DisabledOU -Server $DC -ErrorAction Stop
            $actions += "Moved to Disabled Accounts OU"
            Write-Log "Moved $SamAccountName to $($Global:Config.DisabledOU)"
        } else {
            $actions += "Already in Disabled Accounts OU"
            Write-Log "$SamAccountName already in Disabled OU -- move skipped" "WARN"
        }

        $result.Action = $actions -join "; "
        $result.Status = "SUCCESS"

    } catch {
        $result.Status       = "FAILED"
        $result.ErrorMessage = $_.Exception.Message
        Write-Log "Error processing ${SamAccountName}: $_" "ERROR"
    }

    return $result
}

# ---------------------------------------------------------------------------
# HTML REPORT BUILDER
# ---------------------------------------------------------------------------
function Build-HTMLReport {
    param([array]$Results)

    $successCount = ($Results | Where-Object { $_.Status -eq "SUCCESS" }).Count
    $failCount    = ($Results | Where-Object { $_.Status -eq "FAILED" }).Count
    $totalCount   = $Results.Count
    $runDate      = Get-Date -Format "dddd, MMMM dd yyyy HH:mm"
    $runTime      = Get-Date -Format "HH:mm"
    $dcUsed       = if ($Results.Count -gt 0) { [string]$Results[0].DCUsed } else { "N/A" }
    $logPath      = $Global:Config.LogFile
    $dc1          = $Global:Config.PrimaryDC
    $dc2          = $Global:Config.SecondaryDC

    # Build rows without any special characters
    $rowsSB = New-Object System.Text.StringBuilder
    foreach ($r in $Results) {
        if ($r.Status -eq "SUCCESS") {
            $badge   = '<span style="background:#16a34a;color:#fff;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:600;">SUCCESS</span>'
            $errText = '<span style="color:#16a34a;">OK</span>'
        } else {
            $badge   = '<span style="background:#dc2626;color:#fff;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:600;">FAILED</span>'
            # Escape HTML special chars manually (no System.Web dependency)
            $errEsc  = ([string]$r.ErrorMessage).Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;")
            $errText = "<span style=""color:#dc2626;"">$errEsc</span>"
        }
        $dispName  = ([string]$r.DisplayName).Replace("&","&amp;")
        $samName   = ([string]$r.SamAccountName).Replace("&","&amp;")
        $expDate   = $r.ExpirationDate.ToString('yyyy-MM-dd')
        $actionEsc = ([string]$r.Action).Replace("&","&amp;").Replace("<","&lt;").Replace(">","&gt;")
        $ts        = $r.Timestamp.ToString('HH:mm:ss')

        $null = $rowsSB.Append("<tr>")
        $null = $rowsSB.Append("<td><strong>$samName</strong></td>")
        $null = $rowsSB.Append("<td>$dispName</td>")
        $null = $rowsSB.Append("<td>$expDate</td>")
        $null = $rowsSB.Append("<td style=""font-size:11px;color:#475569;"">$actionEsc</td>")
        $null = $rowsSB.Append("<td>$badge</td>")
        $null = $rowsSB.Append("<td style=""font-size:11px;"">$errText</td>")
        $null = $rowsSB.Append("<td style=""font-size:11px;color:#6b7280;"">$dcUsed</td>")
        $null = $rowsSB.Append("<td style=""font-size:11px;color:#6b7280;"">$ts</td>")
        $null = $rowsSB.Append("</tr>")
    }
    $rowsHtml = $rowsSB.ToString()

    $html = "<!DOCTYPE html><html lang=""en""><head><meta charset=""UTF-8""><style>"
    $html += "* { margin:0; padding:0; box-sizing:border-box; }"
    $html += "body { font-family:'Segoe UI',Arial,sans-serif; background:#f1f5f9; color:#1e293b; }"
    $html += ".wrapper { max-width:920px; margin:30px auto; border-radius:12px; overflow:hidden; box-shadow:0 4px 24px rgba(0,0,0,.12); }"
    $html += ".hdr { background:#1e3a5f; color:#fff; padding:28px 32px; }"
    $html += ".hdr h1 { font-size:20px; font-weight:700; }"
    $html += ".hdr p { font-size:13px; opacity:.8; margin-top:4px; }"
    $html += ".bdg { display:inline-block; background:rgba(255,255,255,.15); border:1px solid rgba(255,255,255,.25); border-radius:20px; padding:3px 12px; font-size:11px; margin-top:8px; margin-right:6px; }"
    $html += ".stats { display:flex; background:#fff; border-bottom:1px solid #e2e8f0; }"
    $html += ".stat { flex:1; padding:16px; text-align:center; border-right:1px solid #e2e8f0; }"
    $html += ".stat:last-child { border-right:none; }"
    $html += ".stat .n { font-size:28px; font-weight:700; line-height:1; }"
    $html += ".stat .l { font-size:11px; color:#64748b; margin-top:4px; text-transform:uppercase; letter-spacing:.5px; }"
    $html += ".info { background:#eff6ff; border-left:4px solid #2563eb; padding:12px 18px; font-size:13px; color:#1e40af; }"
    $html += ".body { background:#fff; padding:22px 28px; }"
    $html += ".body h2 { font-size:14px; font-weight:600; color:#0f172a; padding-bottom:10px; border-bottom:2px solid #e2e8f0; margin-bottom:16px; }"
    $html += "table { width:100%; border-collapse:collapse; font-size:13px; }"
    $html += "thead tr { background:#f8fafc; }"
    $html += "th { padding:9px 11px; text-align:left; font-weight:600; font-size:10px; text-transform:uppercase; letter-spacing:.5px; color:#475569; border-bottom:2px solid #e2e8f0; }"
    $html += "td { padding:9px 11px; border-bottom:1px solid #f1f5f9; vertical-align:middle; }"
    $html += "tr:last-child td { border-bottom:none; }"
    $html += ".ftr { background:#f8fafc; border-top:1px solid #e2e8f0; padding:14px 28px; font-size:11px; color:#94a3b8; text-align:center; }"
    $html += ".ftr strong { color:#64748b; }"
    $html += "</style></head><body><div class=""wrapper"">"

    $html += "<div class=""hdr""><h1>AD Account Expiry Manager - Processing Report</h1>"
    $html += "<p>Business Exception Account Lifecycle Automation</p>"
    $html += "<span class=""bdg"">$runDate</span><span class=""bdg"">DC: $dcUsed</span><span class=""bdg"">Scheduled Run</span></div>"

    $html += "<div class=""stats"">"
    $html += "<div class=""stat""><div class=""n"" style=""color:#2563eb;"">$totalCount</div><div class=""l"">Processed</div></div>"
    $html += "<div class=""stat""><div class=""n"" style=""color:#16a34a;"">$successCount</div><div class=""l"">Succeeded</div></div>"
    $html += "<div class=""stat""><div class=""n"" style=""color:#dc2626;"">$failCount</div><div class=""l"">Failed</div></div>"
    $html += "<div class=""stat""><div class=""n"" style=""color:#7c3aed;"">$runTime</div><div class=""l"">Run Time</div></div>"
    $html += "</div>"

    $html += "<div class=""info"">The accounts below have reached their business exception expiry date. "
    $html += "Each account has been <strong>disabled</strong>, had its <strong>expiration date set</strong>, "
    $html += "and been <strong>moved to the Disabled Accounts OU</strong> to maintain O365 licensing compliance.</div>"

    $html += "<div class=""body""><h2>Account Processing Details</h2>"
    $html += "<table><thead><tr>"
    $html += "<th>Username</th><th>Display Name</th><th>Expiry Date</th><th>Actions Taken</th>"
    $html += "<th>Status</th><th>Error</th><th>DC Used</th><th>Time</th>"
    $html += "</tr></thead><tbody>"
    $html += $rowsHtml
    $html += "</tbody></table></div>"

    $html += "<div class=""ftr""><strong>AD Account Expiry Manager</strong> &middot; "
    $html += "Nodes: <strong>$dc1</strong> / <strong>$dc2</strong> &middot; "
    $html += "Log: <strong>$logPath</strong></div>"
    $html += "</div></body></html>"

    return $html
}

function Save-HTMLReport {
    param([array]$Results)
    try {
        if (-not (Test-Path $Global:Config.ReportDir)) {
            New-Item -ItemType Directory -Path $Global:Config.ReportDir -Force | Out-Null
        }
        $fileName   = "ExpiryReport_$(Get-Date -f 'yyyy-MM-dd_HHmmss').html"
        $reportPath = Join-Path $Global:Config.ReportDir $fileName
        $html       = Build-HTMLReport -Results $Results
        [System.IO.File]::WriteAllText($reportPath, $html, [System.Text.Encoding]::UTF8)
        Write-Log "HTML report saved: $reportPath"
        return $reportPath
    } catch {
        Write-Log "Failed to save HTML report: $_" "ERROR"
        return $null
    }
}

function Send-HTMLReport {
    param([array]$Results)

    $reportPath = Save-HTMLReport -Results $Results

    $successCount = ($Results | Where-Object { $_.Status -eq "SUCCESS" }).Count
    $failCount    = ($Results | Where-Object { $_.Status -eq "FAILED" }).Count
    $subject      = "AD Expiry Manager | $successCount OK, $failCount Failed - $(Get-Date -f 'yyyy-MM-dd')"
    $html         = Build-HTMLReport -Results $Results

    # Validate config before attempting send
    if ($Global:Config.SMTPServer -match "yourdomain") {
        Write-Log "SMTP not configured -- skipping email (SMTPServer is still placeholder)" "WARN"
        Write-Log "Report saved locally: $reportPath"
        return $reportPath
    }

    Write-Log "Attempting SMTP send to $($Global:Config.SMTPServer):$($Global:Config.SMTPPort)..."

    # Safely read optional config keys with fallbacks
    $smtpSSL  = $false
    if ($Global:Config.ContainsKey("SMTPUseSSL")) { $smtpSSL = [bool]$Global:Config.SMTPUseSSL }

    $smtpUser = ""
    if ($Global:Config.ContainsKey("SMTPUser"))   { $smtpUser = [string]$Global:Config.SMTPUser }

    $smtpPass = ""
    if ($Global:Config.ContainsKey("SMTPPass"))   { $smtpPass = [string]$Global:Config.SMTPPass }

    try {
        $smtp                = New-Object System.Net.Mail.SmtpClient
        $smtp.Host           = [string]$Global:Config.SMTPServer
        $smtp.Port           = [int]$Global:Config.SMTPPort
        $smtp.EnableSsl      = $smtpSSL
        $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
        $smtp.Timeout        = 15000

        if ($smtpUser -ne "") {
            $smtp.UseDefaultCredentials = $false
            $smtp.Credentials = New-Object System.Net.NetworkCredential($smtpUser, $smtpPass)
            Write-Log "SMTP using explicit credentials for: $smtpUser"
        } else {
            $smtp.UseDefaultCredentials = $false
            Write-Log "SMTP using anonymous relay mode"
        }

        $msg                 = New-Object System.Net.Mail.MailMessage
        $msg.From            = New-Object System.Net.Mail.MailAddress([string]$Global:Config.SMTPFrom)
        $msg.Subject         = $subject
        $msg.Body            = $html
        $msg.IsBodyHtml      = $true
        $msg.BodyEncoding    = [System.Text.Encoding]::UTF8
        $msg.SubjectEncoding = [System.Text.Encoding]::UTF8

        foreach ($recipient in $Global:Config.ReportRecipients) {
            $r = $recipient.Trim()
            if ($r -ne "") { $msg.To.Add($r) }
        }

        Write-Log "Sending to: $($msg.To -join ', ')"
        $smtp.Send($msg)
        $msg.Dispose()
        $smtp.Dispose()

        Write-Log "HTML report emailed successfully to: $($Global:Config.ReportRecipients -join ', ')"

    } catch [System.Net.Sockets.SocketException] {
        Write-Log "SMTP FAILED [SocketException] -- Cannot reach $($Global:Config.SMTPServer):$($Global:Config.SMTPPort)" "ERROR"
        Write-Log "  Detail: $($_.Exception.Message)" "ERROR"
        Write-Log "  Check: server name, port number, firewall rules" "ERROR"
    } catch [System.Net.Mail.SmtpException] {
        Write-Log "SMTP FAILED [SmtpException] -- Status: $($_.Exception.StatusCode)" "ERROR"
        Write-Log "  Detail: $($_.Exception.Message)" "ERROR"
        Write-Log "  Check: relay permissions, From address, authentication" "ERROR"
    } catch {
        Write-Log "SMTP FAILED [$($_.Exception.GetType().Name)]: $($_.Exception.Message)" "ERROR"
    }

    return $reportPath
}

function Test-SMTPConnection {
    param(
        [string]$Server,
        [int]$Port,
        [string]$From,
        [string]$To,
        [bool]$UseSSL,
        [string]$User,
        [string]$Pass
    )

    $results = @()

    # Step 1: DNS resolve
    try {
        $resolved = [System.Net.Dns]::GetHostAddresses($Server)
        $results += "[OK]   DNS resolved '$Server' -> $($resolved[0].IPAddressToString)"
    } catch {
        $results += "[FAIL] DNS resolution failed for '$Server': $_"
        return $results
    }

    # Step 2: TCP connect
    try {
        $tcp = New-Object System.Net.Sockets.TcpClient
        $ar  = $tcp.BeginConnect($Server, $Port, $null, $null)
        $ok  = $ar.AsyncWaitHandle.WaitOne(3000, $false)
        if ($ok -and $tcp.Connected) {
            $results += "[OK]   TCP port $Port is reachable on $Server"
            $tcp.Close()
        } else {
            $tcp.Close()
            $results += "[FAIL] TCP port $Port is NOT reachable on $Server (timeout)"
            return $results
        }
    } catch {
        $results += "[FAIL] TCP connect to ${Server}:${Port} failed: $_"
        return $results
    }

    # Step 3: Send test email
    try {
        $smtp                = New-Object System.Net.Mail.SmtpClient
        $smtp.Host           = $Server
        $smtp.Port           = $Port
        $smtp.EnableSsl      = $UseSSL
        $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
        $smtp.Timeout        = 10000

        if ($User -and $User -ne "") {
            $smtp.UseDefaultCredentials = $false
            $smtp.Credentials = New-Object System.Net.NetworkCredential($User, $Pass)
            $results += "[INFO] Using credentials: $User"
        } else {
            $smtp.UseDefaultCredentials = $false
            $results += "[INFO] Using anonymous relay mode"
        }

        $msg             = New-Object System.Net.Mail.MailMessage
        $msg.From        = New-Object System.Net.Mail.MailAddress($From)
        $msg.Subject     = "AD Expiry Manager -- SMTP Test $(Get-Date -f 'yyyy-MM-dd HH:mm')"
        $msg.Body        = "Test email from AD Account Expiry Manager.`n`nIf you received this, SMTP is configured correctly.`n`nServer : $Server`nPort   : $Port`nSSL    : $UseSSL`nFrom   : $From`nTo     : $To"
        $msg.IsBodyHtml  = $false
        if ($To.Trim() -ne "") { $msg.To.Add($To.Trim()) }

        $smtp.Send($msg)
        $msg.Dispose()
        $smtp.Dispose()

        $results += "[OK]   Test email sent successfully to $To"
        $results += "[OK]   SMTP configuration is working!"

    } catch [System.Net.Sockets.SocketException] {
        $results += "[FAIL] Socket error: $($_.Exception.Message)"
        $results += "       Check: server name, port, firewall"
    } catch [System.Net.Mail.SmtpException] {
        $results += "[FAIL] SMTP rejected: $($_.Exception.StatusCode) -- $($_.Exception.Message)"
        $results += "       Check: relay permissions, From address, credentials"
    } catch {
        $results += "[FAIL] $($_.Exception.GetType().Name): $($_.Exception.Message)"
    }

    return $results
}

function Send-FailureAlert {
    param([string]$Message)
    if ($Global:Config.SMTPServer -match "yourdomain") { return }
    try {
        $body = "ALERT: $Message`n`nTime: $(Get-Date)`nCheck DCs: $($Global:Config.PrimaryDC), $($Global:Config.SecondaryDC)"

        $smtpSSL = $false
        if ($Global:Config.ContainsKey("SMTPUseSSL")) { $smtpSSL = [bool]$Global:Config.SMTPUseSSL }
        $smtpUser = ""; if ($Global:Config.ContainsKey("SMTPUser")) { $smtpUser = [string]$Global:Config.SMTPUser }
        $smtpPass = ""; if ($Global:Config.ContainsKey("SMTPPass")) { $smtpPass = [string]$Global:Config.SMTPPass }

        $smtp                = New-Object System.Net.Mail.SmtpClient
        $smtp.Host           = [string]$Global:Config.SMTPServer
        $smtp.Port           = [int]$Global:Config.SMTPPort
        $smtp.EnableSsl      = $smtpSSL
        $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
        $smtp.Timeout        = 10000
        if ($smtpUser -ne "") {
            $smtp.UseDefaultCredentials = $false
            $smtp.Credentials = New-Object System.Net.NetworkCredential($smtpUser, $smtpPass)
        } else {
            $smtp.UseDefaultCredentials = $false
        }

        $msg            = New-Object System.Net.Mail.MailMessage
        $msg.From       = New-Object System.Net.Mail.MailAddress([string]$Global:Config.SMTPFrom)
        $msg.Subject    = "AD Expiry Manager -- DC Unreachable Alert"
        $msg.Body       = $body
        $msg.IsBodyHtml = $false
        foreach ($r in $Global:Config.ReportRecipients) {
            if ($r.Trim() -ne "") { $msg.To.Add($r.Trim()) }
        }
        $smtp.Send($msg)
        $msg.Dispose()
        $smtp.Dispose()
    } catch { Write-Log "Failure alert email failed: $($_.Exception.Message)" "WARN" }
}

# ---------------------------------------------------------------------------
# SCHEDULED PROCESSING
# ---------------------------------------------------------------------------
function Invoke-ScheduledProcessing {
    Write-Log "=== Scheduled processing started ==="

    $accounts = @(Load-AccountList)
    $today    = (Get-Date).Date
    $results  = @()

    try { $dc = Get-AvailableDC }
    catch {
        Write-Log $_ "ERROR"
        Send-FailureAlert -Message $_
        return $null
    }

    foreach ($entry in $accounts) {
        if (-not $entry.ExpirationDate) { continue }
        $expDate = $null
        try { $expDate = [datetime]$entry.ExpirationDate } catch { continue }

        if ($expDate.Date -le $today -and $entry.Status -ne "Disabled") {
            Write-Log "Processing: $($entry.SamAccountName) (expiry: $($expDate.ToString('yyyy-MM-dd')))"
            $result              = Disable-ExceptionAccount -SamAccountName $entry.SamAccountName `
                                                            -ExpirationDate $expDate -DC $dc
            $results            += $result
            $entry.Status        = if ($result.Status -eq "SUCCESS") { "Disabled" } else { "Error: $($result.ErrorMessage)" }
            $entry.LastProcessed = Get-Date
        }
    }

    Save-AccountList -List $accounts

    if ($results.Count -gt 0) {
        Write-Log "Processed $($results.Count) account(s). Building report..."
        $reportPath = Send-HTMLReport -Results $results
        Write-Log "=== Scheduled processing completed. Results: $($results.Count) ==="
        return $reportPath
    } else {
        Write-Log "No accounts due for processing today."
        Write-Log "=== Scheduled processing completed. No actions taken. ==="
        return $null
    }
}

# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
function Show-GUI {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    # Use a module-level variable for the accounts list so all
    # event handlers share the same reference without scope issues
    $Global:AccountList = @(Load-AccountList)

    # ── Main form ────────────────────────────────────────────────────────────
    $form                 = New-Object System.Windows.Forms.Form
    $form.Text            = "AD Account Expiry Manager  --  Business Exception Handler"
    $form.Size            = New-Object System.Drawing.Size(1020, 700)
    $form.StartPosition   = "CenterScreen"
    $form.BackColor       = [System.Drawing.Color]::FromArgb(241, 245, 249)
    $form.Font            = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.FormBorderStyle = "FixedSingle"
    $form.MaximizeBox     = $false

    # ── Header ───────────────────────────────────────────────────────────────
    $headerPanel           = New-Object System.Windows.Forms.Panel
    $headerPanel.Location  = New-Object System.Drawing.Point(0, 0)
    $headerPanel.Size      = New-Object System.Drawing.Size(1020, 68)
    $headerPanel.BackColor = [System.Drawing.Color]::FromArgb(30, 58, 95)
    $form.Controls.Add($headerPanel)

    $lblTitle           = New-Object System.Windows.Forms.Label
    $lblTitle.Text      = "AD Account Expiry Manager"
    $lblTitle.Font      = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $lblTitle.ForeColor = [System.Drawing.Color]::White
    $lblTitle.Location  = New-Object System.Drawing.Point(20, 10)
    $lblTitle.Size      = New-Object System.Drawing.Size(550, 28)
    $headerPanel.Controls.Add($lblTitle)

    $lblSub           = New-Object System.Windows.Forms.Label
    $lblSub.Text      = "Business Exception Handler  |  Nodes: Contoso-dcsrv-01 / Contoso-dcsrv-02"
    $lblSub.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $lblSub.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184)
    $lblSub.Location  = New-Object System.Drawing.Point(22, 40)
    $lblSub.Size      = New-Object System.Drawing.Size(500, 20)
    $headerPanel.Controls.Add($lblSub)

    $lblDC1           = New-Object System.Windows.Forms.Label
    $lblDC1.Text      = "[ ] Contoso-dcsrv-01  Checking..."
    $lblDC1.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184)
    $lblDC1.Location  = New-Object System.Drawing.Point(730, 18)
    $lblDC1.Size      = New-Object System.Drawing.Size(200, 20)
    $headerPanel.Controls.Add($lblDC1)

    $lblDC2           = New-Object System.Windows.Forms.Label
    $lblDC2.Text      = "[ ] Contoso-dcsrv-02  Checking..."
    $lblDC2.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184)
    $lblDC2.Location  = New-Object System.Drawing.Point(730, 40)
    $lblDC2.Size      = New-Object System.Drawing.Size(200, 20)
    $headerPanel.Controls.Add($lblDC2)

    $form.Add_Shown({
        foreach ($dc in @($Global:Config.PrimaryDC, $Global:Config.SecondaryDC)) {
            $lbl = if ($dc -eq $Global:Config.PrimaryDC) { $lblDC1 } else { $lblDC2 }
            if (Test-Connection -ComputerName $dc -Count 1 -Quiet -ErrorAction SilentlyContinue) {
                $lbl.ForeColor = [System.Drawing.Color]::FromArgb(74, 222, 128)
                $lbl.Text      = "[Online]  $dc"
            } else {
                $lbl.ForeColor = [System.Drawing.Color]::FromArgb(248, 113, 113)
                $lbl.Text      = "[Offline] $dc"
            }
        }
    })

    # ── Add Account section ───────────────────────────────────────────────────
    $gbAdd           = New-Object System.Windows.Forms.GroupBox
    $gbAdd.Text      = "  Add Business Exception Account"
    $gbAdd.Location  = New-Object System.Drawing.Point(14, 80)
    $gbAdd.Size      = New-Object System.Drawing.Size(990, 110)
    $gbAdd.ForeColor = [System.Drawing.Color]::FromArgb(30, 58, 95)
    $gbAdd.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($gbAdd)

    $lblUser           = New-Object System.Windows.Forms.Label
    $lblUser.Text      = "AD Username (sAMAccountName):"
    $lblUser.Location  = New-Object System.Drawing.Point(14, 30)
    $lblUser.Size      = New-Object System.Drawing.Size(210, 20)
    $lblUser.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
    $lblUser.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $gbAdd.Controls.Add($lblUser)

    $txtUser             = New-Object System.Windows.Forms.TextBox
    $txtUser.Location    = New-Object System.Drawing.Point(14, 52)
    $txtUser.Size        = New-Object System.Drawing.Size(200, 28)
    $txtUser.Font        = New-Object System.Drawing.Font("Segoe UI", 10)
    $txtUser.BorderStyle = "FixedSingle"
    $gbAdd.Controls.Add($txtUser)

    $lblExp           = New-Object System.Windows.Forms.Label
    $lblExp.Text      = "Account Expiration Date:"
    $lblExp.Location  = New-Object System.Drawing.Point(236, 30)
    $lblExp.Size      = New-Object System.Drawing.Size(180, 20)
    $lblExp.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
    $lblExp.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $gbAdd.Controls.Add($lblExp)

    $dtpExp              = New-Object System.Windows.Forms.DateTimePicker
    $dtpExp.Location     = New-Object System.Drawing.Point(236, 52)
    $dtpExp.Size         = New-Object System.Drawing.Size(180, 28)
    $dtpExp.Format       = "Custom"
    $dtpExp.CustomFormat = "yyyy-MM-dd"
    $dtpExp.MinDate      = (Get-Date).AddDays(1)
    $dtpExp.Value        = (Get-Date).AddDays(30)
    $dtpExp.Font         = New-Object System.Drawing.Font("Segoe UI", 10)
    $gbAdd.Controls.Add($dtpExp)

    $lblNotes           = New-Object System.Windows.Forms.Label
    $lblNotes.Text      = "Business Justification (optional):"
    $lblNotes.Location  = New-Object System.Drawing.Point(440, 30)
    $lblNotes.Size      = New-Object System.Drawing.Size(220, 20)
    $lblNotes.ForeColor = [System.Drawing.Color]::FromArgb(51, 65, 85)
    $lblNotes.Font      = New-Object System.Drawing.Font("Segoe UI", 9)
    $gbAdd.Controls.Add($lblNotes)

    $txtNotes            = New-Object System.Windows.Forms.TextBox
    $txtNotes.Location   = New-Object System.Drawing.Point(440, 52)
    $txtNotes.Size       = New-Object System.Drawing.Size(300, 28)
    $txtNotes.Font       = New-Object System.Drawing.Font("Segoe UI", 10)
    $txtNotes.BorderStyle = "FixedSingle"
    $gbAdd.Controls.Add($txtNotes)

    $btnVerify                            = New-Object System.Windows.Forms.Button
    $btnVerify.Text                       = "Verify User"
    $btnVerify.Location                   = New-Object System.Drawing.Point(755, 46)
    $btnVerify.Size                       = New-Object System.Drawing.Size(100, 36)
    $btnVerify.BackColor                  = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $btnVerify.ForeColor                  = [System.Drawing.Color]::White
    $btnVerify.FlatStyle                  = "Flat"
    $btnVerify.FlatAppearance.BorderSize  = 0
    $btnVerify.Font                       = New-Object System.Drawing.Font("Segoe UI", 9)
    $btnVerify.Cursor                     = "Hand"
    $gbAdd.Controls.Add($btnVerify)

    $btnAdd                               = New-Object System.Windows.Forms.Button
    $btnAdd.Text                          = "+ Add Account"
    $btnAdd.Location                      = New-Object System.Drawing.Point(864, 46)
    $btnAdd.Size                          = New-Object System.Drawing.Size(116, 36)
    $btnAdd.BackColor                     = [System.Drawing.Color]::FromArgb(37, 99, 235)
    $btnAdd.ForeColor                     = [System.Drawing.Color]::White
    $btnAdd.FlatStyle                     = "Flat"
    $btnAdd.FlatAppearance.BorderSize     = 0
    $btnAdd.Font                          = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnAdd.Cursor                        = "Hand"
    $gbAdd.Controls.Add($btnAdd)

    # ── Account List section ──────────────────────────────────────────────────
    $gbList           = New-Object System.Windows.Forms.GroupBox
    $gbList.Text      = "  Exception Account Queue"
    $gbList.Location  = New-Object System.Drawing.Point(14, 200)
    $gbList.Size      = New-Object System.Drawing.Size(990, 360)
    $gbList.ForeColor = [System.Drawing.Color]::FromArgb(30, 58, 95)
    $gbList.Font      = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $form.Controls.Add($gbList)

    $lvAccounts              = New-Object System.Windows.Forms.ListView
    $lvAccounts.Location     = New-Object System.Drawing.Point(10, 22)
    $lvAccounts.Size         = New-Object System.Drawing.Size(968, 290)
    $lvAccounts.View         = "Details"
    $lvAccounts.FullRowSelect = $true
    $lvAccounts.GridLines    = $true
    $lvAccounts.Font         = New-Object System.Drawing.Font("Segoe UI", 9)
    $lvAccounts.BorderStyle  = "FixedSingle"
    $gbList.Controls.Add($lvAccounts)

    @(
        @("Username",       140),
        @("Display Name",   180),
        @("Expiry Date",    100),
        @("Days Remaining", 100),
        @("Status",         130),
        @("Added By",       110),
        @("Added On",       130),
        @("Justification",  162)
    ) | ForEach-Object {
        $col       = New-Object System.Windows.Forms.ColumnHeader
        $col.Text  = $_[0]
        $col.Width = [int]$_[1]
        $lvAccounts.Columns.Add($col) | Out-Null
    }

    $btnRemove                            = New-Object System.Windows.Forms.Button
    $btnRemove.Text                       = "Remove Selected"
    $btnRemove.Location                   = New-Object System.Drawing.Point(10, 322)
    $btnRemove.Size                       = New-Object System.Drawing.Size(140, 32)
    $btnRemove.BackColor                  = [System.Drawing.Color]::FromArgb(220, 38, 38)
    $btnRemove.ForeColor                  = [System.Drawing.Color]::White
    $btnRemove.FlatStyle                  = "Flat"
    $btnRemove.FlatAppearance.BorderSize  = 0
    $btnRemove.Font                       = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnRemove.Cursor                     = "Hand"
    $gbList.Controls.Add($btnRemove)

    $btnRefresh                           = New-Object System.Windows.Forms.Button
    $btnRefresh.Text                      = "Refresh"
    $btnRefresh.Location                  = New-Object System.Drawing.Point(160, 322)
    $btnRefresh.Size                      = New-Object System.Drawing.Size(80, 32)
    $btnRefresh.BackColor                 = [System.Drawing.Color]::FromArgb(71, 85, 105)
    $btnRefresh.ForeColor                 = [System.Drawing.Color]::White
    $btnRefresh.FlatStyle                 = "Flat"
    $btnRefresh.FlatAppearance.BorderSize = 0
    $btnRefresh.Cursor                    = "Hand"
    $gbList.Controls.Add($btnRefresh)

    $btnOpenReport                           = New-Object System.Windows.Forms.Button
    $btnOpenReport.Text                      = "Open Last Report"
    $btnOpenReport.Location                  = New-Object System.Drawing.Point(250, 322)
    $btnOpenReport.Size                      = New-Object System.Drawing.Size(130, 32)
    $btnOpenReport.BackColor                 = [System.Drawing.Color]::FromArgb(124, 58, 237)
    $btnOpenReport.ForeColor                 = [System.Drawing.Color]::White
    $btnOpenReport.FlatStyle                 = "Flat"
    $btnOpenReport.FlatAppearance.BorderSize = 0
    $btnOpenReport.Cursor                    = "Hand"
    $gbList.Controls.Add($btnOpenReport)

    $btnSMTP                                 = New-Object System.Windows.Forms.Button
    $btnSMTP.Text                            = "SMTP Settings & Test"
    $btnSMTP.Location                        = New-Object System.Drawing.Point(390, 322)
    $btnSMTP.Size                            = New-Object System.Drawing.Size(155, 32)
    $btnSMTP.BackColor                       = [System.Drawing.Color]::FromArgb(15, 118, 110)
    $btnSMTP.ForeColor                       = [System.Drawing.Color]::White
    $btnSMTP.FlatStyle                       = "Flat"
    $btnSMTP.FlatAppearance.BorderSize       = 0
    $btnSMTP.Cursor                          = "Hand"
    $gbList.Controls.Add($btnSMTP)

    $btnRunNow                            = New-Object System.Windows.Forms.Button
    $btnRunNow.Text                       = "Run Processing Now"
    $btnRunNow.Location                   = New-Object System.Drawing.Point(798, 322)
    $btnRunNow.Size                       = New-Object System.Drawing.Size(162, 32)
    $btnRunNow.BackColor                  = [System.Drawing.Color]::FromArgb(5, 150, 105)
    $btnRunNow.ForeColor                  = [System.Drawing.Color]::White
    $btnRunNow.FlatStyle                  = "Flat"
    $btnRunNow.FlatAppearance.BorderSize  = 0
    $btnRunNow.Font                       = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnRunNow.Cursor                     = "Hand"
    $gbList.Controls.Add($btnRunNow)

    # ── Status bar ────────────────────────────────────────────────────────────
    $statusBar           = New-Object System.Windows.Forms.Label
    $statusBar.Location  = New-Object System.Drawing.Point(0, 630)
    $statusBar.Size      = New-Object System.Drawing.Size(1020, 26)
    $statusBar.BackColor = [System.Drawing.Color]::FromArgb(30, 41, 59)
    $statusBar.ForeColor = [System.Drawing.Color]::FromArgb(148, 163, 184)
    $statusBar.Text      = "  Ready  |  Data: $($Global:Config.DataFile)"
    $statusBar.Font      = New-Object System.Drawing.Font("Segoe UI", 8)
    $statusBar.TextAlign = "MiddleLeft"
    $form.Controls.Add($statusBar)

    # ── Helper functions (defined before event handlers) ──────────────────────
    function Set-Status {
        param([string]$msg, [string]$color = "Gray")
        $statusBar.Text = "  $msg"
        $statusBar.ForeColor = switch ($color) {
            "Green"  { [System.Drawing.Color]::FromArgb(74,  222, 128) }
            "Red"    { [System.Drawing.Color]::FromArgb(248, 113, 113) }
            "Yellow" { [System.Drawing.Color]::FromArgb(250, 204,  21) }
            default  { [System.Drawing.Color]::FromArgb(148, 163, 184) }
        }
        $form.Refresh()
    }

    function Refresh-ListView {
        $lvAccounts.Items.Clear()
        $today = (Get-Date).Date

        foreach ($a in $Global:AccountList) {
            if (-not $a.ExpirationDate) { continue }
            $expiry = $null
            try { $expiry = [datetime]$a.ExpirationDate } catch { continue }

            $days = [math]::Ceiling(($expiry - $today).TotalDays)
            $daysText = ""
            if    ($days -lt 0) { $daysText = "Expired" }
            elseif ($days -eq 0) { $daysText = "Today" }
            else  { $daysText = "$days days" }

            $fDisplayName   = if ($a.DisplayName)   { [string]$a.DisplayName }   else { "" }
            $fStatus        = if ($a.Status)         { [string]$a.Status }        else { "Pending" }
            $fAddedBy       = if ($a.AddedBy)        { [string]$a.AddedBy }       else { "" }
            $fAddedOn       = if ($a.AddedOn)        { [string]$a.AddedOn }       else { "" }
            $fJustification = if ($a.Justification)  { [string]$a.Justification } else { "" }

            $item = New-Object System.Windows.Forms.ListViewItem($a.SamAccountName)
            $item.SubItems.Add($fDisplayName)                  | Out-Null
            $item.SubItems.Add($expiry.ToString('yyyy-MM-dd')) | Out-Null
            $item.SubItems.Add($daysText)                      | Out-Null
            $item.SubItems.Add($fStatus)                       | Out-Null
            $item.SubItems.Add($fAddedBy)                      | Out-Null
            $item.SubItems.Add($fAddedOn)                      | Out-Null
            $item.SubItems.Add($fJustification)                | Out-Null

            if ($fStatus -eq "Disabled") {
                $item.BackColor = [System.Drawing.Color]::FromArgb(240, 253, 244)
                $item.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
            } elseif ($days -le 0) {
                $item.BackColor = [System.Drawing.Color]::FromArgb(254, 242, 242)
                $item.ForeColor = [System.Drawing.Color]::FromArgb(153, 27, 27)
            } elseif ($days -le 7) {
                $item.BackColor = [System.Drawing.Color]::FromArgb(255, 251, 235)
                $item.ForeColor = [System.Drawing.Color]::FromArgb(146, 64, 14)
            }
            $lvAccounts.Items.Add($item) | Out-Null
        }

        $due = 0
        foreach ($a in $Global:AccountList) {
            if ($a.Status -ne "Disabled" -and $a.ExpirationDate) {
                $exp = $null
                try { $exp = [datetime]$a.ExpirationDate } catch { }
                if ($exp -and $exp -le (Get-Date)) { $due++ }
            }
        }

        $statusColor = if ($due -gt 0) { "Yellow" } else { "Gray" }
        Set-Status "Queue: $($Global:AccountList.Count) account(s)  |  Due/overdue: $due  |  $(Get-Date -f 'HH:mm:ss')" $statusColor
    }

    # Initial load
    Refresh-ListView

    # ── Event handlers ────────────────────────────────────────────────────────

    # Verify
    $btnVerify.Add_Click({
        $sam = $txtUser.Text.Trim()
        if (-not $sam) {
            [System.Windows.Forms.MessageBox]::Show("Enter a username first.", "Validation", "OK", "Warning") | Out-Null
            return
        }
        Set-Status "Verifying '$sam' in AD..." "Yellow"
        try {
            $dc   = Get-AvailableDC
            $user = Get-ADUserSafe -SamAccountName $sam -DC $dc
            if ($user) {
                $info = "User found on $dc`n`nDisplay Name : $($user.DisplayName)`nDepartment   : $($user.Department)`nTitle        : $($user.Title)`nEnabled      : $($user.Enabled)`nEmail        : $($user.EmailAddress)"
                [System.Windows.Forms.MessageBox]::Show($info, "User Verified", "OK", "Information") | Out-Null
                Set-Status "User '$sam' verified OK." "Green"
            } else {
                [System.Windows.Forms.MessageBox]::Show("User '$sam' was not found in Active Directory.", "Not Found", "OK", "Warning") | Out-Null
                Set-Status "User '$sam' not found." "Red"
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("AD connection error:`n$_", "Error", "OK", "Error") | Out-Null
            Set-Status "AD connection error." "Red"
        }
    })

    # Add Account
    $btnAdd.Add_Click({
        $sam    = $txtUser.Text.Trim()
        $expiry = $dtpExp.Value.Date
        $notes  = $txtNotes.Text.Trim()

        if (-not $sam) {
            [System.Windows.Forms.MessageBox]::Show("Please enter an AD username.", "Validation", "OK", "Warning") | Out-Null
            return
        }
        if ($expiry -le (Get-Date).Date) {
            [System.Windows.Forms.MessageBox]::Show("Expiration date must be in the future.", "Validation", "OK", "Warning") | Out-Null
            return
        }

        # Check duplicate
        $duplicate = $false
        foreach ($a in $Global:AccountList) {
            if ($a.SamAccountName -eq $sam) { $duplicate = $true; break }
        }
        if ($duplicate) {
            [System.Windows.Forms.MessageBox]::Show("'$sam' is already in the exception queue.", "Duplicate", "OK", "Warning") | Out-Null
            return
        }

        Set-Status "Looking up display name for '$sam'..." "Yellow"
        $displayName = $sam
        try {
            $dc   = Get-AvailableDC
            $user = Get-ADUserSafe -SamAccountName $sam -DC $dc
            if ($user -and $user.DisplayName) { $displayName = [string]$user.DisplayName }
        } catch { }

        $entry = [PSCustomObject]@{
            SamAccountName = $sam
            DisplayName    = $displayName
            ExpirationDate = $expiry
            Status         = "Pending"
            Justification  = $notes
            AddedBy        = $env:USERNAME
            AddedOn        = (Get-Date -f 'yyyy-MM-dd HH:mm')
            LastProcessed  = $null
        }

        # Append to global list and save
        $Global:AccountList = @($Global:AccountList) + @($entry)
        Save-AccountList -List $Global:AccountList

        $txtUser.Clear()
        $txtNotes.Clear()
        $dtpExp.Value = (Get-Date).AddDays(30)

        Refresh-ListView
        Set-Status "Account '$sam' added. Total: $($Global:AccountList.Count)" "Green"
        Write-Log "Account added: $sam | Expiry: $($expiry.ToString('yyyy-MM-dd')) | By: $($env:USERNAME)"
    })

    # Remove
    $btnRemove.Add_Click({
        if ($lvAccounts.SelectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Select an account to remove.", "No Selection", "OK", "Warning") | Out-Null
            return
        }
        $sam     = $lvAccounts.SelectedItems[0].Text
        $confirm = [System.Windows.Forms.MessageBox]::Show(
            "Remove '$sam' from the queue?`n`nNote: This does NOT re-enable the AD account.",
            "Confirm Remove", "YesNo", "Question")
        if ($confirm -eq "Yes") {
            $newList = @()
            foreach ($a in $Global:AccountList) {
                if ($a.SamAccountName -ne $sam) { $newList += $a }
            }
            $Global:AccountList = $newList
            Save-AccountList -List $Global:AccountList
            Refresh-ListView
            Set-Status "Account '$sam' removed. Total: $($Global:AccountList.Count)" "Yellow"
            Write-Log "Account removed: $sam by $($env:USERNAME)"
        }
    })

    # Refresh
    $btnRefresh.Add_Click({
        $Global:AccountList = @(Load-AccountList)
        Refresh-ListView
        Set-Status "List refreshed from disk. Total: $($Global:AccountList.Count)" "Gray"
    })

    # Open Last Report
    $btnOpenReport.Add_Click({
        $reportDir = $Global:Config.ReportDir
        if (Test-Path $reportDir) {
            $latest = Get-ChildItem -Path $reportDir -Filter "*.html" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($latest) {
                Start-Process $latest.FullName
                Set-Status "Opened report: $($latest.Name)" "Green"
            } else {
                [System.Windows.Forms.MessageBox]::Show("No reports found in:`n$reportDir", "No Reports", "OK", "Information") | Out-Null
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show("Reports folder does not exist yet.`nRun processing first.", "No Reports", "OK", "Information") | Out-Null
        }
    })

    # Run Now
    $btnRunNow.Add_Click({
        $dueCount = 0
        foreach ($a in $Global:AccountList) {
            if ($a.Status -ne "Disabled" -and $a.ExpirationDate) {
                $exp = $null
                try { $exp = [datetime]$a.ExpirationDate } catch { }
                if ($exp -and $exp.Date -le (Get-Date).Date) { $dueCount++ }
            }
        }

        if ($dueCount -eq 0) {
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "No accounts are due today.`n`nDo you want to run a processing check anyway?",
                "Nothing Due", "YesNo", "Question")
            if ($confirm -ne "Yes") { return }
        } else {
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                "$dueCount account(s) are due for processing.`n`nEach will be:`n- Disabled in Active Directory`n- Expiration date applied`n- Moved to Disabled Accounts OU`n- Included in HTML report`n`nContinue?",
                "Confirm Processing", "YesNo", "Question")
            if ($confirm -ne "Yes") { return }
        }

        Set-Status "Processing -- please wait..." "Yellow"
        $form.Enabled = $false
        try {
            $reportPath = Invoke-ScheduledProcessing
            $Global:AccountList = @(Load-AccountList)
            Refresh-ListView

            if ($reportPath -and (Test-Path $reportPath)) {
                $openReport = [System.Windows.Forms.MessageBox]::Show(
                    "Processing complete!`n`nHTML report saved to:`n$reportPath`n`nOpen the report now?",
                    "Done", "YesNo", "Information")
                if ($openReport -eq "Yes") { Start-Process $reportPath }
                Set-Status "Processing complete. Report: $reportPath" "Green"
            } else {
                [System.Windows.Forms.MessageBox]::Show(
                    "Processing complete. No accounts were due, so no report was generated.",
                    "Done", "OK", "Information") | Out-Null
                Set-Status "Processing complete. No report generated." "Gray"
            }
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error during processing:`n$_", "Error", "OK", "Error") | Out-Null
            Set-Status "Processing error -- check log." "Red"
        } finally {
            $form.Enabled = $true
        }
    })

    # SMTP Settings & Test
    $btnSMTP.Add_Click({

        # ---- Build SMTP dialog ----
        $dlg                 = New-Object System.Windows.Forms.Form
        $dlg.Text            = "SMTP Settings & Connection Test"
        $dlg.Size            = New-Object System.Drawing.Size(540, 560)
        $dlg.StartPosition   = "CenterParent"
        $dlg.BackColor       = [System.Drawing.Color]::FromArgb(248, 250, 252)
        $dlg.Font            = New-Object System.Drawing.Font("Segoe UI", 9)
        $dlg.FormBorderStyle = "FixedDialog"
        $dlg.MaximizeBox     = $false
        $dlg.MinimizeBox     = $false

        $pnlTop              = New-Object System.Windows.Forms.Panel
        $pnlTop.Location     = New-Object System.Drawing.Point(0, 0)
        $pnlTop.Size         = New-Object System.Drawing.Size(540, 50)
        $pnlTop.BackColor    = [System.Drawing.Color]::FromArgb(15, 118, 110)
        $dlg.Controls.Add($pnlTop)
        $lblH           = New-Object System.Windows.Forms.Label
        $lblH.Text      = "SMTP Configuration & Test"
        $lblH.Font      = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
        $lblH.ForeColor = [System.Drawing.Color]::White
        $lblH.Location  = New-Object System.Drawing.Point(16, 14)
        $lblH.Size      = New-Object System.Drawing.Size(500, 24)
        $pnlTop.Controls.Add($lblH)

        function Add-DlgLabel { param($text,$x,$y)
            $l = New-Object System.Windows.Forms.Label
            $l.Text = $text; $l.Location = New-Object System.Drawing.Point($x,$y)
            $l.Size = New-Object System.Drawing.Size(200,18)
            $l.ForeColor = [System.Drawing.Color]::FromArgb(51,65,85)
            $dlg.Controls.Add($l); return $l }

        function Add-DlgText { param($x,$y,$w,$val)
            $t = New-Object System.Windows.Forms.TextBox
            $t.Location = New-Object System.Drawing.Point($x,$y)
            $t.Size = New-Object System.Drawing.Size($w,24)
            $t.Text = $val; $t.BorderStyle = "FixedSingle"
            $dlg.Controls.Add($t); return $t }

        Add-DlgLabel "SMTP Server:"          16  68  | Out-Null
        $tServer  = Add-DlgText 16  88  260 $Global:Config.SMTPServer

        Add-DlgLabel "Port:"                 290 68  | Out-Null
        $tPort    = Add-DlgText 290 88  60  ([string]$Global:Config.SMTPPort)

        $chkSSL           = New-Object System.Windows.Forms.CheckBox
        $chkSSL.Text      = "Use SSL/TLS"
        $chkSSL.Location  = New-Object System.Drawing.Point(370, 90)
        $chkSSL.Size      = New-Object System.Drawing.Size(100, 20)
        $chkSSL.Checked   = [bool]$Global:Config.SMTPUseSSL
        $dlg.Controls.Add($chkSSL)

        Add-DlgLabel "From Address:"         16  120 | Out-Null
        $tFrom    = Add-DlgText 16  140 340 $Global:Config.SMTPFrom

        Add-DlgLabel "Username (leave blank for relay):" 16 170 | Out-Null
        $tUser    = Add-DlgText 16  190 340 $Global:Config.SMTPUser

        Add-DlgLabel "Password (leave blank for relay):" 16 220 | Out-Null
        $tPass    = Add-DlgText 16  240 340 $Global:Config.SMTPPass
        $tPass.PasswordChar = [char]"*"[0]

        Add-DlgLabel "Test Recipient:"       16  270 | Out-Null
        $tTestTo  = Add-DlgText 16  290 340 ($Global:Config.ReportRecipients[0])

        Add-DlgLabel "Report Recipients (one per line):" 16 320 | Out-Null
        $tRecip               = New-Object System.Windows.Forms.TextBox
        $tRecip.Location      = New-Object System.Drawing.Point(16, 340)
        $tRecip.Size          = New-Object System.Drawing.Size(490, 60)
        $tRecip.Multiline     = $true
        $tRecip.ScrollBars    = "Vertical"
        $tRecip.BorderStyle   = "FixedSingle"
        $tRecip.Text          = ($Global:Config.ReportRecipients -join "`r`n")
        $dlg.Controls.Add($tRecip)

        $lblResult           = New-Object System.Windows.Forms.Label
        $lblResult.Location  = New-Object System.Drawing.Point(16, 410)
        $lblResult.Size      = New-Object System.Drawing.Size(490, 16)
        $lblResult.ForeColor = [System.Drawing.Color]::FromArgb(71,85,105)
        $lblResult.Text      = ""
        $dlg.Controls.Add($lblResult)

        $txtLog              = New-Object System.Windows.Forms.TextBox
        $txtLog.Location     = New-Object System.Drawing.Point(16, 428)
        $txtLog.Size         = New-Object System.Drawing.Size(490, 60)
        $txtLog.Multiline    = $true
        $txtLog.ReadOnly     = $true
        $txtLog.ScrollBars   = "Vertical"
        $txtLog.BorderStyle  = "FixedSingle"
        $txtLog.BackColor    = [System.Drawing.Color]::FromArgb(15, 23, 42)
        $txtLog.ForeColor    = [System.Drawing.Color]::FromArgb(74, 222, 128)
        $txtLog.Font         = New-Object System.Drawing.Font("Consolas", 8)
        $dlg.Controls.Add($txtLog)

        $btnTest                            = New-Object System.Windows.Forms.Button
        $btnTest.Text                       = "Send Test Email"
        $btnTest.Location                   = New-Object System.Drawing.Point(16, 498)
        $btnTest.Size                       = New-Object System.Drawing.Size(140, 30)
        $btnTest.BackColor                  = [System.Drawing.Color]::FromArgb(15, 118, 110)
        $btnTest.ForeColor                  = [System.Drawing.Color]::White
        $btnTest.FlatStyle                  = "Flat"
        $btnTest.FlatAppearance.BorderSize  = 0
        $btnTest.Font                       = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
        $dlg.Controls.Add($btnTest)

        $btnSave                            = New-Object System.Windows.Forms.Button
        $btnSave.Text                       = "Save Settings"
        $btnSave.Location                   = New-Object System.Drawing.Point(166, 498)
        $btnSave.Size                       = New-Object System.Drawing.Size(120, 30)
        $btnSave.BackColor                  = [System.Drawing.Color]::FromArgb(37, 99, 235)
        $btnSave.ForeColor                  = [System.Drawing.Color]::White
        $btnSave.FlatStyle                  = "Flat"
        $btnSave.FlatAppearance.BorderSize  = 0
        $dlg.Controls.Add($btnSave)

        $btnClose                           = New-Object System.Windows.Forms.Button
        $btnClose.Text                      = "Close"
        $btnClose.Location                  = New-Object System.Drawing.Point(296, 498)
        $btnClose.Size                      = New-Object System.Drawing.Size(80, 30)
        $btnClose.BackColor                 = [System.Drawing.Color]::FromArgb(71, 85, 105)
        $btnClose.ForeColor                 = [System.Drawing.Color]::White
        $btnClose.FlatStyle                 = "Flat"
        $btnClose.FlatAppearance.BorderSize = 0
        $dlg.Controls.Add($btnClose)

        $btnTest.Add_Click({
            $txtLog.Clear()
            $txtLog.ForeColor = [System.Drawing.Color]::FromArgb(74, 222, 128)
            $lblResult.Text   = "Testing SMTP connection..."
            $dlg.Refresh()

            $portNum = 25
            try { $portNum = [int]$tPort.Text } catch { }

            $lines = Test-SMTPConnection `
                -Server  $tServer.Text.Trim() `
                -Port    $portNum `
                -From    $tFrom.Text.Trim() `
                -To      $tTestTo.Text.Trim() `
                -UseSSL  $chkSSL.Checked `
                -User    $tUser.Text.Trim() `
                -Pass    $tPass.Text

            $txtLog.Lines = $lines
            $success = ($lines | Where-Object { $_ -like "*[OK]*SMTP configuration*" }).Count -gt 0
            if ($success) {
                $lblResult.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
                $lblResult.Text      = "Test email sent successfully!"
            } else {
                $lblResult.ForeColor = [System.Drawing.Color]::FromArgb(153, 27, 27)
                $lblResult.Text      = "SMTP test failed -- see details above"
                $txtLog.ForeColor    = [System.Drawing.Color]::FromArgb(248, 113, 113)
            }
        })

        $btnSave.Add_Click({
            $portNum = 25
            try { $portNum = [int]$tPort.Text } catch { }
            $recipList = @($tRecip.Text -split "`r`n|`n" | Where-Object { $_.Trim() -ne "" } | ForEach-Object { $_.Trim() })

            $Global:Config.SMTPServer       = $tServer.Text.Trim()
            $Global:Config.SMTPPort         = $portNum
            $Global:Config.SMTPUseSSL       = $chkSSL.Checked
            $Global:Config.SMTPFrom         = $tFrom.Text.Trim()
            $Global:Config.SMTPUser         = $tUser.Text.Trim()
            $Global:Config.SMTPPass         = $tPass.Text
            $Global:Config.ReportRecipients = $recipList

            [System.Windows.Forms.MessageBox]::Show(
                "Settings saved for this session.`n`nTo make permanent, update the Config block`nat the top of the script file.",
                "Saved", "OK", "Information") | Out-Null
            $lblResult.ForeColor = [System.Drawing.Color]::FromArgb(22, 101, 52)
            $lblResult.Text      = "Settings applied. Remember to update the script file."
            Set-Status "SMTP settings updated: $($tServer.Text.Trim()):$portNum" "Green"
        })

        $btnClose.Add_Click({ $dlg.Close() })
        $dlg.ShowDialog($form) | Out-Null
    })

    [System.Windows.Forms.Application]::Run($form)

    # Cleanup global when form closes
    Remove-Variable -Name AccountList -Scope Global -ErrorAction SilentlyContinue
}

# ---------------------------------------------------------------------------
# ENTRY POINT
# ---------------------------------------------------------------------------
if ($args -contains "-Scheduled") {
    Invoke-ScheduledProcessing
} else {
    Show-GUI
}
