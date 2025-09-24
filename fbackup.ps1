param (
    [Parameter(Mandatory = $true)] [string]$TaskName,
    [switch]$TestMail,
    [string]$ffsLogFolder = "$env:USERPROFILE\AppData\Roaming\FreeFileSync\Logs"
)

# Find FreeFileSync.exe path
function Get-FreeFileSyncInstallPath {
    $regPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($regPath in $regPaths) {
        $ffsInstall = Get-ItemProperty -Path $regPath -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like "*FreeFileSync*" -and $_.InstallLocation } |
            Select-Object -First 1

        if ($ffsInstall) {
            return $ffsInstall.InstallLocation
        }
    }
    return $null
}

# Load XML config
function Load-Config {
    if (Test-Path -LiteralPath $xmlConfigFile) {
        return [xml](Get-Content -LiteralPath $xmlConfigFile)
    }
    return $null
}

# Interactive wizard to create XML config
function Run-Wizard {
    Write-Host "`n--- Configuration Wizard ---`n"

    Write-Host "Press ENTER to accept the value shown in [brackets]."
    Write-Host "Valid options are shown in (parentheses).`n"

    $config = New-Object System.Xml.XmlDocument
    $root   = $config.CreateElement("Config")
    $null   = $config.AppendChild($root)

    $questions = @(
        @{Name="SendEmail";    Default="Everytime"; Options=@("Everytime","Never","OnlyError")},
        @{Name="SmtpServer";   Default="smtp.example.com"},
        @{Name="Port";         Default="465"},
        @{Name="UseSSL";       Default="false";     Options=@("true","false")},
        @{Name="User";         Default=""},
        @{Name="Password";     Default=""},
        @{Name="From";         Default=""},
        @{Name="To";           Default=""},
        @{Name="SubjectPrefix";Default=""}
    )

    # Collect user input
    foreach ($q in $questions) {
        $prompt = "$($q.Name) [$($q.Default)]"
        if ($q.Options) { $prompt += " (" + ($q.Options -join "/") + ")" }

        do {
            $response = Read-Host $prompt
            if ([string]::IsNullOrWhiteSpace($response)) { 
                $response = $q.Default 
            }
            if ($q.Options -and ($response -notin $q.Options)) {
                Write-Host "Invalid value. Valid options: $($q.Options -join ', ')"
                continue
            }
            break
        } while ($true)

        # Encrypt password
        if ($q.Name -eq "Password") {
            $secure   = ConvertTo-SecureString $response -AsPlainText -Force
            $response = ConvertFrom-SecureString $secure
        }

        # Suggest From = User
        if ($q.Name -eq "User" -and $response) {
            ($questions | Where-Object Name -eq "From").Default = $response
        }

        $elem = $config.CreateElement($q.Name)
        $elem.InnerText = $response
        $null = $root.AppendChild($elem)
    }

    # Save config?
    do {
        $save = (Read-Host "`nSave configuration? (y/n)").Trim().ToLower()
    } until ($save -in @("y","n","yes","no"))

    if ($save -in @("y","yes")) {
        $config.Save($xmlConfigFile)
        Write-Host "Configuration saved to $xmlConfigFile"
    } else {
        Write-Host "[WARNING] Configuration not saved."
    }

    # Test email?
    do {
        $test = (Read-Host "`nTest email now? (y/n)").Trim().ToLower()
    } until ($test -in @("y","n","yes","no"))

    if ($test -in @("y","yes")) {
        $sent = Send-TestEmail $config
        if ($sent) {
            Write-Host "Test email sent successfully."
        } else {
            Write-Host "[ERROR] Test email failed."
        }
    }

    Write-Host "`nConfiguration completed."
    Write-Host "Run the script again to execute FreeFileSync."
    exit
}

# Send email
function Send-Email {
    param (
        [xml]$Config,
        [string]$Subject,
        [string]$Body,
        [string[]]$AttachmentPaths = @()
    )

    $smtp          = $Config.Config.SmtpServer
    $port          = [int]$Config.Config.Port
    $useSSL        = $Config.Config.UseSSL.ToLower() -eq 'true'
    $from          = $Config.Config.From
    $to            = ($Config.Config.To -split '[;,]' | ForEach-Object { $_.Trim() }) -ne ''
    $user          = $Config.Config.User
    $encryptedPass = $Config.Config.Password
    $prefix        = $Config.Config.SubjectPrefix
    
    try {
        $tcpClient   = New-Object System.Net.Sockets.TcpClient
        $asyncResult = $tcpClient.BeginConnect($smtp, $port, $null, $null)
        if (-not $asyncResult.AsyncWaitHandle.WaitOne(5000, $false)) {
            throw "Connection timeout after 5 seconds"
        }
        $tcpClient.EndConnect($asyncResult)
        Write-Host "[MAIL] Connection to $($smtp):$($port) succeeded."
        $tcpClient.Close()
    } catch {
        Write-Host "[MAIL] ERROR: SMTP connection failed: $($_.Exception.Message)"
        return $null
    }

    # Decrypt password
    try {
        $securePass = ConvertTo-SecureString $encryptedPass
    } catch {
        Write-Host "[MAIL] ERROR: Failed to decrypt password: $($_.Exception.Message)"
        return $null
    }

    $cred = New-Object System.Management.Automation.PSCredential($user, $securePass)

    $mailParams = @{
        From       = $from
        To         = $to
        Subject    = "$($prefix)$($Subject)"
        Body       = $Body
        SmtpServer = $smtp
        Port       = $port
        Credential = $cred
        UseSsl     = $useSSL
        DeliveryNotificationOption = "OnFailure"
        Priority   = "Normal"
    }

    if ($AttachmentPaths.Count -gt 0) {
        $existing = $AttachmentPaths | Where-Object { Test-Path $_ }
        if ($existing.Count -eq 0) {
            Write-Host "[MAIL] WARNING: No valid attachments found, skipping."
        } else {
            $mailParams.Add("Attachments", $existing)
            Write-Host "[MAIL] Attachments: $($existing -join ', ')"
        }
    }

    try {
        Write-Host "[MAIL] Sending message..."
        Send-MailMessage @mailParams -ErrorAction Stop
        return $true
    } catch {
        Write-Host "[MAIL] ERROR: Email sending failed: $($_.Exception.Message)"
        if ($useSSL -and $port -eq 465) {
            Write-Host "[MAIL] Trying fallback TLS to $($smtp):587..."
            $mailParams.UseSsl = $false
            $mailParams.Port = 587
            try {
                Send-MailMessage @mailParams -ErrorAction Stop
                Write-Host "[MAIL] Email sent via TLS fallback."
                return $true
            } catch {
                Write-Host "[MAIL] ERROR: TLS fallback failed: $($_.Exception.Message)"
            }
        }
    }

    return $null
}

# Send test email
function Send-TestEmail {
    param ([xml]$Config)
    Send-Email -Config $Config -Subject "Test Email" -Body "This is a test email from fBackup script."
}

# Execute FreeFileSync
function Execute-FFS {
    #3..1 | ForEach-Object { Write-Host "$_" ; Start-Sleep -Seconds 1 }
    $proc = Start-Process -FilePath "$ffsPath\FreeFileSync.exe" -ArgumentList "`"$ffsBatchFile`"" -Wait -PassThru
    $exitCode = $proc.ExitCode    
    return $exitCode
}

# Get latest log file and zip it
function Get-FFS-LogFile {
    param(
        [datetime]$StartTime,
        [int]$MaxRetries = 5,
        [int]$DelaySeconds = 1
    )    

    $latest = $null
    for ($i = 1; $i -le $MaxRetries -and -not $latest; $i++) {
        $latest = Get-ChildItem -LiteralPath $ffsLogFolder -File -ErrorAction SilentlyContinue |
            Where-Object {
                $_.Name.StartsWith($TaskName, [System.StringComparison]::OrdinalIgnoreCase) -and
                $_.LastWriteTime -ge $StartTime
            } |
            Sort-Object LastWriteTime -Descending |
            Select-Object -First 1

        if (-not $latest) {
            Write-Host "[DEBUG] Attempt $($i)/$($MaxRetries): no log found yet, retrying in $($DelaySeconds) sec..."
            Start-Sleep -Seconds $DelaySeconds
        }
    }

    if (-not $latest) {
        Write-Host "[WARNING] No log file found in '$($ffsLogFolder)' for task '$($TaskName)' after $($StartTime)."
        return $null
    }

    $safe = ($latest.BaseName -replace '[^a-zA-Z0-9]', '_')
    if ([string]::IsNullOrWhiteSpace($safe)) {
        $safe = "log_$([System.Guid]::NewGuid().ToString('N'))"
    }
    $zipPath = Join-Path $env:TEMP "$($safe).zip"

    try {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop

        $fs = [System.IO.File]::Open($zipPath, [System.IO.FileMode]::CreateNew)
        try {
            $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create, $false)
            try {
                $src = [System.IO.File]::Open($latest.FullName, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read)
                try {
                    $entry = $zip.CreateEntry($latest.Name, [System.IO.Compression.CompressionLevel]::Optimal)
                    $src.CopyTo($entry.Open())
                } finally { $src.Dispose() }
            } finally { $zip.Dispose() }
        } finally { $fs.Dispose() }
    } catch {
        Write-Host "[WARNING] Failed to create zip archive: $($_.Exception.Message)"
        return $null
    }

    return $zipPath
}



# --- Main ---

# Globals
$ffsBatchFile  = "$($TaskName).ffs_batch"
$xmlConfigFile = "$($TaskName).xml"
$ffsPath       = Get-FreeFileSyncInstallPath

# Check ffsPath
if (-not $ffsPath) {
    Write-Host "[ERROR] FreeFileSync not installed."
    exit 1 
}

# Check batch file
if (-not (Test-Path $ffsBatchFile)) {
    Write-Host "[ERROR] FreeFileSync task '$($TaskName)' not found. Expected batch file '$($ffsBatchFile)'."
    exit 1
}

# Check config file
$config = Load-Config
if (-not $config) {
    Run-Wizard
    Exit 0
}

# Test email mode
if ($TestMail) {
    if (Send-TestEmail $config) {
        Write-Host "Test email sent successfully."
    } else {
        Write-Host "[ERROR] Test email failed."
    }
    exit
}

# Run FreeFileSync
Write-Host "Executing FreeFileSync..."
$startTime = Get-Date #to ensure logfile is newer than execution
$exitCode  = Execute-FFS
Write-Host "FreeFileSync finished with exit code $($exitCode)"

# Decide whether to send email
$sendMode   = $config.Config.SendEmail.ToLower()
$shouldSend = $false
switch ($sendMode) {
    "everytime"   { $shouldSend = $true }
    "onlyerror"   { if ($exitCode -ne 0) { $shouldSend = $true } }
    "never"       { $shouldSend = $false }
    default       { Write-Host "[WARNING] Unknown SendEmail mode '$sendMode'. Email will not be sent." }
}

# Send email if required
if ($shouldSend) {
    Write-Host "Preparing report email for task '$($TaskName)'..."

    # Build subject
    $subject = "$($SubjectPrefix) - $($TaskName): $($result)"

    # Build body
    switch ($exitCode) {
        0 { $body = "FreeFileSync finished with no errors." }
        1 { $body = "FreeFileSync finished with warnings." }
        2 { $body = "FreeFileSync finished with errors." }
        default { $body = "FreeFileSync finished with unknown result." }
    }

    # Attach log only if warning or error
    $attachments = @()
    if ($exitCode -ne 0) {
        $zipLog = Get-FFS-LogFile -StartTime $startTime
        if (-not $zipLog) {
            Write-Host "[WARNING] No log file found to attach."
            $body += " (no log available)"
        } else {
            $attachments = @($zipLog)
        }
    }

    $sent = Send-Email -Config $config -Subject $subject -Body $body -AttachmentPath $attachments
    if ($sent) {
        # Remove zipLog
        if ($zipLog) {
            $removed = $false
            for ($i=0; $i -lt 3 -and -not $removed; $i++) {
                try {
                    Remove-Item -Path $zipLog -ErrorAction Stop
                    Write-Host "Temporary log file removed."
                    $removed = $true
                } catch {
                    Write-Host "[WARNING] Attempt $($i+1) to remove temp log file failed: $($_.Exception.Message)"                    
                    Start-Sleep -Seconds 5
                }
            }
            if (-not $removed) {
                Write-Host "[WARNING] Could not remove temp log file '$zipLog'."
            }
        }
        Write-Host "Email sent successfully."
    } else {
        Write-Host "[ERROR] Report email failed"
    }
} else {
    Write-Host "Email not required (mode: $($sendMode))."
}