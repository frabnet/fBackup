 param (
    [Parameter(Mandatory=$True)]
    [string]$nomeTask
 ) 

# -------------------------------------------------- Funzioni 

#FIXME: Aggiornare la funzione con ultima versione https://github.com/lipkau/PsIni

function Get-IniContent {  
    <#  
    .Synopsis  
        Gets the content of an INI file  
          
    .Description  
        Gets the content of an INI file and returns it as a hashtable  
          
    .Notes  
        Author        : Oliver Lipkau <oliver@lipkau.net>  
        Blog        : http://oliver.lipkau.net/blog/  
        Source        : https://github.com/lipkau/PsIni 
                      http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91 
        Version        : 1.0 - 2010/03/12 - Initial release  
                      1.1 - 2014/12/11 - Typo (Thx SLDR) 
                                         Typo (Thx Dave Stiff) 
          
        #Requires -Version 2.0  
          
    .Inputs  
        System.String  
          
    .Outputs  
        System.Collections.Hashtable  
          
    .Parameter FilePath  
        Specifies the path to the input file.  
          
    .Example  
        $FileContent = Get-IniContent "C:\myinifile.ini"  
        -----------  
        Description  
        Saves the content of the c:\myinifile.ini in a hashtable called $FileContent  
      
    .Example  
        $inifilepath | $FileContent = Get-IniContent  
        -----------  
        Description  
        Gets the content of the ini file passed through the pipe into a hashtable called $FileContent  
      
    .Example  
        C:\PS>$FileContent = Get-IniContent "c:\settings.ini"  
        C:\PS>$FileContent["Section"]["Key"]  
        -----------  
        Description  
        Returns the key "Key" of the section "Section" from the C:\settings.ini file  
          
    .Link  
        Out-IniFile  
    #>  
      
    [CmdletBinding()]  
    Param(  
        [ValidateNotNullOrEmpty()]  
        [ValidateScript({(Test-Path $_) -and ((Get-Item $_).Extension -eq ".ini")})]  
        [Parameter(ValueFromPipeline=$True,Mandatory=$True)]  
        [string]$FilePath  
    )  
      
    Begin  
        {Write-Verbose "$($MyInvocation.MyCommand.Name):: Function started"}  
          
    Process  
    {  
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Processing file: $Filepath"  
              
        $ini = @{}  
        switch -regex -file $FilePath  
        {  
            "^\[(.+)\]$" # Section  
            {  
                $section = $matches[1]  
                $ini[$section] = @{}  
                $CommentCount = 0  
            }  
            "^(;.*)$" # Comment  
            {  
                if (!($section))  
                {  
                    $section = "No-Section"  
                    $ini[$section] = @{}  
                }  
                $value = $matches[1]  
                $CommentCount = $CommentCount + 1  
                $name = "Comment" + $CommentCount  
                $ini[$section][$name] = $value  
            }   
            "(.+?)\s*=\s*(.*)" # Key  
            {  
                if (!($section))  
                {  
                    $section = "No-Section"  
                    $ini[$section] = @{}  
                }  
                $name,$value = $matches[1..2]  
                $ini[$section][$name] = $value  
            }  
        }  
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Finished Processing file: $FilePath"  
        Return $ini  
    }  
          
    End  
        {Write-Verbose "$($MyInvocation.MyCommand.Name):: Function ended"}  
} 

function Espelli {    
    param ( [string]$label ) 
    Write-Host -NoNewLine "Espulsione disco $label... "
    $vol = Get-WmiObject Win32_Volume -filter "Label = '$label'"
    if ($vol -ne $null) {
        $Eject =  New-Object -comObject Shell.Application
        $Eject.NameSpace(17).ParseName($vol.driveletter).InvokeVerb("Eject")
        Write-Host "Fatto"
    } else {
        Write-Host "Unita non trovata"
    }
}

function InvioEmail {
    param( [string]$from, [string]$to, [string]$subject, [string]$body,[string]$attachmentPath,[string]$server,[string]$serverPort,[string]$enableSSL )

    #Non funzionava con Aruba che utilizza Implicit SSL
    #fix: http://nicholasarmstrong.com/2009/12/sending-email-with-powershell-implicit-and-explicit-ssl/

    #Controllo file Credenziali
    If (Test-Path $emailCredentialsFile) {
        $emailCredentials = Import-CliXml $emailCredentialsFile
    } else {
        $emailCredentials = Get-Credential -Message "Inserire credenziali per $emailSmtpServer" | Export-CliXml $emailCredentialsFile
        $emailCredentials = Import-CliXml $emailCredentialsFile
    }
    $credentials = [Net.NetworkCredential]($emailCredentials)

    
    if ($enableSSL -eq '0') {
        # Set up server connection
        $smtpClient = New-Object System.Net.Mail.SmtpClient $server, $serverPort
        $smtpClient.EnableSsl = $false
        $smtpClient.Timeout = $timeout
        $smtpClient.UseDefaultCredentials = $false;
        $smtpClient.Credentials = $credentials

        $message = New-Object System.Net.Mail.MailMessage $from, $to, $subject, $body

        # Allegato
        if ($attachmentPath -ne '') {
            $attachment = New-Object System.Net.Mail.Attachment $attachmentPath
            $message.Attachments.Add($attachment)            
        }

        # Send the message
        Write-Host -NoNewline "Invio email a $to... "
        try
        {
            $smtpClient.Send($message)
            Write-Output "Inviato."
        }
        catch
        {
            Write-Error $_
            Write-Output "ERRORE."
        }
    } else {
        # Load System.Web assembly
        [System.Reflection.Assembly]::LoadWithPartialName("System.Web") > $null

        # Create a new mail with the appropriate server settigns
        $mail = New-Object System.Web.Mail.MailMessage
        $mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserver", $server)
        $mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpserverport", $serverPort)
        $mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpusessl", $true)
        $mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", $credentials.UserName)
        $mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", $credentials.Password)
        $mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout", $timeout / 1000)
        # Use network SMTP server...
        $mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusing", 2)
        # ... and basic authentication
        $mail.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", 1)

        # Set up the mail message fields
        $mail.From = $from
        $mail.To = $to
        $mail.Subject = $subject
        $mail.Body = $body

        # Allegato
        if ($attachmentPath -ne '') {
            # Convert to full path and attach file to message
            $attachmentPath = (get-item $attachmentPath).FullName
            $attachment = New-Object System.Web.Mail.MailAttachment $attachmentPath
            $mail.Attachments.Add($attachment) > $null
        }

        # Send the message
        Write-Host -NoNewline "Invio email a $to... "
        try
        {
            [System.Web.Mail.SmtpMail]::Send($mail)
            Write-Host  "Inviato."
        }
        catch
        {
            Write-Error $_
            Write-Host  "ERRORE."
        }
    }
}

function Comprimi {
    param ( [string]$FileName ) 
    Write-Host -NoNewLine "Compressione file di log... "
    If (Test-Path 'logs.zip') { Remove-Item 'logs.zip' }
    $sz_args =  ' a -tzip -mx9 -y logs.zip ' + '"' + $FileName + '"'
    $process = (Start-Process -FilePath "7za.exe" -PassThru -Wait -ArgumentList $sz_args)
    If ( $process.ExitCode -eq 0 ) { Write-Host "Ok!" } else {  Write-Host 'ERRORE' $process.ExitCode }
}

# -------------------------------------------------- Variabili task

#FIXME: $LogGeneral = Join-Path $LogPath 'Gen.log' /PER UNIRE I PERCORSI - Elegante

# Controllo se è stato lanciato correttamente http://ramblingcookiemonster.github.io/Task-Scheduler/

#$CurrentDir = (Get-Item -Path ".\" -Verbose).FullName
#$Date = Get-Date
#$Admin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
#$Whoami = whoami
#"$Date - Admin=$Admin User=$Whoami Cd=$CurrentDir" | Out-File "AdminLog.txt" -Append

$inifile = $nomeTask + '.ini'
$conf = Get-IniContent $inifile

$emailEnable = $conf['email']['Enable']
$emailFrom = $conf['email']['From']
$emailTo = $conf['email']['To']
$emailToCcnErr = $conf['email']['CcnErr']
$emailSubjPrefix = $conf['email']['Subject']
$emailSmtpServer = $conf['email']['SmtpServer']
$emailSmtpPort = $conf['email']['SmtpPort']
$emailSmtpSSL = $conf['email']['EnableSSL']
$usbEjectLabel = $conf['usb']['EjectLabel']
$usbDayToEject = $conf['usb']['EjectDay']
$shutdown = $conf['varie']['shutdown']
$servstop = $conf['varie']['stopservice']

$Folder = "..\Logs"
$emailCredentialsFile = 'email-auth.xml'

$ServiceWasRunning = $false

# -------------------------------------------------- Inizio programma 

#Modalità test email
$emailEnable = $emailEnable.ToUpper()
If ( $emailEnable -eq "TEST" ) {

    Write-Host "--------------------"
    Write-Host "MODALITA' TEST EMAIL"
    Write-Host "--------------------"

    $msgSubject = "Backup [$nomeTask]: Test email"
    $msgBody = @'
______________________________________________________________
|01/01/2016 - TestEMAIL: Sincronizzazione test test
|
|    Oggetti processati: 999.999 (99.9 GB)
|    Tempo totale: 09:09:09
|_____________________________________________________________
'@     
    Write-Host -NoNewLine "Creazione finto file di log... "
    for($i=1; $i -le 1000 ;  $i++){ Add-Content 'TestLog.txt' $msgBody }
    for($i=1; $i -le 2000 ;  $i++){ Get-Random | Add-Content 'TestLog.txt' }
    Write-Host "Ok."
    Comprimi('TestLog.txt')
    Remove-Item('TestLog.txt')
    InvioEmail $emailFrom $emailTo "$msgSubject" "$msgBody" logs.zip $emailSmtpServer $emailSmtpPort $emailSmtpSSL
    
    Read-Host -Prompt "Premere INVIO per uscire."    
    Exit
}


#Stop del servizio
If ($servstop -ne '') {
    $srvObj = Get-Service -displayName $servstop
    If ( $srvObj.Status -eq 'Running' ) {
        Write-Host "Stop servizio  $servstop"
        $ServiceWasRunning = $true
        Stop-Service -displayName $servstop
    }
}

#Esecuzione FreeFileSync
Write-Host -NoNewLine "Esecuzione FreeFileSync... "
$ffs_args =  "fbackup\$nomeTask.ffs_batch"
$process = Start-Process -FilePath "FreeFileSync.exe" -PassThru -Wait -WorkingDirectory ".." -ArgumentList $ffs_args -Verb runAs
$ffsReturnCode = $process.ExitCode

#Oggetto Email
switch ($ffsReturnCode) { 
    3 { $msgSubject = "$emailSubjPrefix [$nomeTask]: Annullato" }
    2 { $msgSubject = "$emailSubjPrefix [$nomeTask]: Errore"}
    1 { $msgSubject = "$emailSubjPrefix [$nomeTask]: Avvisi" }
    0 { $msgSubject = "$emailSubjPrefix [$nomeTask]: Completato con successo" }
    default { $msgSubject = "$emailSubjPrefix [$nomeTask]: Errore" }
}
Write-Host "Fatto ($ffsReturnCode)"

#Start servizio
If ($ServiceWasRunning) {
    Write-Host "Start servizio  $servstop"
    Start-Service -displayName $servstop    
}

#Cerco l'ultimo file di log
Write-Host -NoNewLine "Ricerca nuovo file... "
Get-ChildItem -Path "$Folder" -Filter "$nomeTask*log" | Sort LastWriteTime –Descending | Select -First 1 | Foreach-Object { $LastLogFile = $_.Name }
$LastLogFile = $LastLogFile.Trim()
Write-Host $LastLogFile
$LastLogFile = $Folder + '\' + $LastLogFile

#Creazione corpo messaggio
$Write = 0
$msgBody = ''
$lines = Get-Content -LiteralPath $LastLogFile
ForEach ($line in $lines) {
    $line = $line.Trim()
    if ($line -match "____" ) { $Write = 1 }    
    if ($Write = 1) { $msgBody = $msgBody + $line + "`r`n" }
    if ($line -match "\|____" ) { break }
}

#Logica invio mail - controllo se devo spedire
$toSend = $false
If ($emailEnable -eq 'EVERYTIME') { $toSend = $true }
If ($emailEnable -eq 'ONLYERROR') { if ($ffs_process.ExitCode -ne 0) { $toSend = $true } }

#Logica invio mail - allego solo se errore - spedisco a $emailCcnErr
if ($ffsReturnCode -eq 0) {
    $attachment = ''
} else {
    Comprimi($LastLogFile)
    $attachment = (Resolve-Path .\).Path + '\logs.zip'
    If ($emailToCcnErr -ne '') { InvioEmail $emailFrom $emailToCcnErr "$msgSubject" "$msgBody" $attachment $emailSmtpServer $emailSmtpPort $emailSmtpSSL }
}
if ($toSend) { InvioEmail $emailFrom $emailTo "$msgSubject" "$msgBody" $attachment $emailSmtpServer $emailSmtpPort $emailSmtpSSL }

#Espulsione unità usb
if ($usbDayToEject -eq (Get-Date).DayOfWeek -Or $usbDayToEject -eq 0 ) {
    Espelli($usbEjectLabel)
}

#Shutdown
if ($shutdown -eq 1) {
    shutdown /s /f /t 60
}

#Uscita con errorcode di FreeFileSync
Exit $ffsReturnCode