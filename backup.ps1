 param (
    [Parameter(Mandatory=$True)]
    [string]$nomeTask
 ) 

#https://github.com/lipkau/PsIni/blob/master/PSIni/Functions/Get-IniContent.ps1
Function Get-IniContent {
    <#
    .Synopsis
        Gets the content of an INI file

    .Description
        Gets the content of an INI file and returns it as a hashtable

    .Notes
        Author		: Oliver Lipkau <oliver@lipkau.net>
		Source		: https://github.com/lipkau/PsIni
                      http://gallery.technet.microsoft.com/scriptcenter/ea40c1ef-c856-434b-b8fb-ebd7a76e8d91
        Version		: 1.0.0 - 2010/03/12 - OL - Initial release
                      1.0.1 - 2014/12/11 - OL - Typo (Thx SLDR)
                                              Typo (Thx Dave Stiff)
                      1.0.2 - 2015/06/06 - OL - Improvment to switch (Thx Tallandtree)
                      1.0.3 - 2015/06/18 - OL - Migrate to semantic versioning (GitHub issue#4)
                      1.0.4 - 2015/06/18 - OL - Remove check for .ini extension (GitHub Issue#6)
                      1.1.0 - 2015/07/14 - CB - Improve round-tripping and be a bit more liberal (GitHub Pull #7)
                                           OL - Small Improvments and cleanup
                      1.1.1 - 2015/07/14 - CB - changed .outputs section to be OrderedDictionary
                      1.1.2 - 2016/08/18 - SS - Add some more verbose outputs as the ini is parsed,
                      				            allow non-existent paths for new ini handling,
                      				            test for variable existence using local scope,
                      				            added additional debug output.

        #Requires -Version 2.0

    .Inputs
        System.String

    .Outputs
        System.Collections.Specialized.OrderedDictionary

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
    [OutputType(
        [System.Collections.Specialized.OrderedDictionary]
    )]
    Param(
        # Specifies the path to the input file.
        [ValidateNotNullOrEmpty()]
        [Parameter( Mandatory = $true, ValueFromPipeline = $true )]
        [String]
        $FilePath,

        # Specify what characters should be describe a comment.
        # Lines starting with the characters provided will be rendered as comments.
        # Default: ";"
        [Char[]]
        $CommentChar = @(";"),

        # Remove lines determined to be comments from the resulting dictionary.
        [Switch]
        $IgnoreComments
    )

    Begin {
        Write-Debug "PsBoundParameters:"
        $PSBoundParameters.GetEnumerator() | ForEach-Object { Write-Debug $_ }
        if ($PSBoundParameters['Debug']) {
            $DebugPreference = 'Continue'
        }
        Write-Debug "DebugPreference: $DebugPreference"

        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function started"

        $commentRegex = "^([$($CommentChar -join '')].*)$"

        Write-Debug ("commentRegex is {0}." -f $commentRegex)
    }

    Process {
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Processing file: $Filepath"

        $ini = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::OrdinalIgnoreCase)

        if (!(Test-Path $Filepath)) {
            Write-Verbose ("Warning: `"{0}`" was not found." -f $Filepath)
            Write-Output $ini
        }

        $commentCount = 0
        switch -regex -file $FilePath {
            "^\s*\[(.+)\]\s*$" {
                # Section
                $section = $matches[1]
                Write-Verbose "$($MyInvocation.MyCommand.Name):: Adding section : $section"
                $ini[$section] = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::OrdinalIgnoreCase)
                $CommentCount = 0
                continue
            }
            $commentRegex {
                # Comment
                if (!$IgnoreComments) {
                    if (!(test-path "variable:local:section")) {
                        $section = $script:NoSection
                        $ini[$section] = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::OrdinalIgnoreCase)
                    }
                    $value = $matches[1]
                    $CommentCount++
                    Write-Debug ("Incremented CommentCount is now {0}." -f $CommentCount)
                    $name = "Comment" + $CommentCount
                    Write-Verbose "$($MyInvocation.MyCommand.Name):: Adding $name with value: $value"
                    $ini[$section][$name] = $value
                }
                else {
                    Write-Debug ("Ignoring comment {0}." -f $matches[1])
                }

                continue
            }
            "(.+?)\s*=\s*(.*)" {
                # Key
                if (!(test-path "variable:local:section")) {
                    $section = $script:NoSection
                    $ini[$section] = New-Object System.Collections.Specialized.OrderedDictionary([System.StringComparer]::OrdinalIgnoreCase)
                }
                $name, $value = $matches[1..2]
                Write-Verbose "$($MyInvocation.MyCommand.Name):: Adding key $name with value: $value"
                $ini[$section][$name] = $value
                continue
            }
        }
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Finished Processing file: $FilePath"
        Write-Output $ini
    }

    End {
        Write-Verbose "$($MyInvocation.MyCommand.Name):: Function ended"
    }
}

Function Espelli {    
    param ( [string]$label ) 
    Write-Host -NoNewLine "Espulsione disco $label... "
    $vol = Get-WmiObject Win32_Volume -filter "Label = '$label'"
    if ($vol -ne $null) {
        Start-Process -FilePath "RemoveDrive.exe" -ArgumentList $vol.DriveLetter
        Write-Host "Done"
    } else {
        Write-Host "Label not found"
    }
}

Function Comprimi {
    param ( [string]$FileName ) 
    Write-Host -NoNewLine "Compressing log file... "
    If (Test-Path 'logs.zip') { Remove-Item 'logs.zip' }
    $sz_args =  ' a -tzip -mx9 -y logs.zip ' + '"' + $FileName + '"'
    $process = (Start-Process -FilePath "7za.exe" -PassThru -Wait -ArgumentList $sz_args)
    If ( $process.ExitCode -eq 0 ) { Write-Host "done." } else {  Write-Host 'ERRORE' $process.ExitCode }
}

Function InvioEmail {
    param ( [string]$To, [string]$Subject, [string]$Body, [string]$Attachment ) 

    #Controllo file Credenziali
    If (Test-Path $emailCredentialsFile) {
        $emailCredentials = Import-CliXml $emailCredentialsFile
    } else {
        $emailCredentials = Get-Credential -Message "Please enter login information for $emailSmtpServer" | Export-CliXml $emailCredentialsFile
        $emailCredentials = Import-CliXml $emailCredentialsFile
    }

    #Supporto destinatari multipli
    $To.Split(';') | ForEach {
        $mailsplat = @{
            To=$_
            From=$emailFrom
            Body=$Body
            Subject=$Subject
            SmtpServer=$emailSmtpServer
            Port=$emailSmtpPort            
            Credential = $emailCredentials
        }

        if ($Attachment -ne '') { $mailsplat.Add('Attachments', $Attachment) }
        if ([boolean]$emailSmtpSSL) { $mailsplat.Add('UseSsl', [bool]1) }
        
        Write-Host -NoNewline "Sending email to $_... "
        try {
            [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { return $true }
            Send-MailMessage @mailsplat
            Write-Host "Sent."
        } catch {
            Write-Error $_
            Write-Output "ERROR."
        }


    }
}

# -------------------------------------------------- Lettura parametri

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

$emailEnable = $emailEnable.ToUpper()
$LogFolder = "..\Logs"
$emailCredentialsFile = $nomeTask + '-auth.xml'

# -------------------------------------------------- Inizio programma 

#Modalità test email
If ( $emailEnable -eq "TEST" ) {
    Write-Host "---------------"
    Write-Host "EMAIL TEST MODE"
    Write-Host "---------------"
    $msgSubject = "$emailSubjPrefix [$nomeTask]: Avvisi"
    $msgBody = @'
______________________________________________________________
|01/01/2016 - $nomeTask: Avvisi
|
|    Oggetti processati: 999.999 (99.9 GB)
|    Tempo totale: 09:09:09
|_____________________________________________________________
'@     
    Write-Host -NoNewLine "Creating fake log file... "
    for($i=1; $i -le 100 ;  $i++){ Add-Content 'TestLog.txt' $msgBody }
    for($i=1; $i -le 500 ;  $i++){ Get-Random | Add-Content 'TestLog.txt' }
    Write-Host "done."
    Comprimi('TestLog.txt')
    Remove-Item('TestLog.txt')

    InvioEmail -To $emailTo -Subject "$msgSubject" -Body "$msgBody" -Attachment "logs.zip"
    
    Read-Host -Prompt "Press ENTER to quit"
    Exit
}

#Stop del servizio
If ($servstop -ne '') {
    $srvObj = Get-Service -displayName $servstop
    If ( $srvObj.Status -eq 'Running' ) {
        Write-Host "Stopping service $servstop"
        $ServiceWasRunning = $true
        Stop-Service -displayName $servstop
    }
}

#Esecuzione FreeFileSync
Write-Host -NoNewLine "Executing FreeFileSync... "
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
Write-Host "Done ($ffsReturnCode)"

#Start servizio
If ($ServiceWasRunning) {
    Write-Host "Starting service $servstop"
    Start-Service -displayName $servstop    
}

#Cerco l'ultimo file di log
Write-Host -NoNewLine "Searching for latest logfile... "
Get-ChildItem -Path "$LogFolder" -Filter "$nomeTask*log" | Sort LastWriteTime –Descending | Select -First 1 | Foreach-Object { $LastLogFile = $_.Name }
$LastLogFile = $LastLogFile.Trim()
Write-Host $LastLogFile
$LastLogFile = $LogFolder + '\' + $LastLogFile

#Creazione corpo messaggio -> $msgBody
$Write = 0
$msgBody = ''
$lines = Get-Content -LiteralPath $LastLogFile
ForEach ($line in $lines) {
    $line = $line.Trim()
    if ($line -match "____" ) { $Write = 1 }    
    if ($Write = 1) { $msgBody = $msgBody + $line + "`r`n" }
    if ($line -match "\|____" ) { break }
}

#Logica invio mail - controllo se devo spedire -> $toSend
$toSend = $false
If ($emailEnable -eq 'EVERYTIME') { $toSend = $true }
If ($emailEnable -eq 'ONLYERROR') { if ($ffsReturnCode -ne 0) { $toSend = $true } }

#Logica invio mail - allego solo se errore - spedisco a $emailCcnErr
if ($ffsReturnCode -eq 0) {
    $attachment = ''
} else {
    Comprimi($LastLogFile)
    $attachment = (Resolve-Path .\).Path + '\logs.zip'
    If ($emailToCcnErr -ne '') { InvioEmail -To $emailToCcnErr -Subject "$msgSubject" -Body "$msgBody" -Attachment $attachment }
}

#Invio email
if ($toSend) {    
    InvioEmail -To $emailTo -Subject "$msgSubject" -Body "$msgBody" -Attachment $attachment
}

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