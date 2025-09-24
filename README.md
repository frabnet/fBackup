## fBackup

This tool is a helper for FreeFileSync. It can send you email with FreeFileSync logs and automate some backup related tasks.
I currently use it on 10-20 servers and constantly try to improve/fix bugs.

## ⚠️ Upgrade Notice (v2)

This release is a full rewrite of fBackup.  
Previous configuration files are **not compatible** and must be recreated.

- Old files `TaskName.ini` and `TaskName-auth.xml` are no longer used.  
- All configuration is now stored in a single XML file: `TaskName.xml`.  
- Options `Shutdown PC` and `Eject USB` have been removed (they were rarely used).  

After upgrading, please re-run `run_fbackup.cmd` for each task to recreate the configuration.

## Installation

1. Download and install FreeFileSync: [Download the Latest Version - FreeFileSync](https://www.freefilesync.org/download.php)  
   Suggested install path: `C:\FreeFileSync\`

2. Download fBackup directly from GitHub [GitHub - frabnet/fBackup](https://github.com/frabnet/fBackup). -> Code -> Download ZIP

3. Unzip `fBackup` inside the FreeFileSync folder:  
   `C:\FreeFileSync\fBackup`

## Task Configuration

1. Create a new Job with FreeFileSync as you like.

2. Save it as a Batch Job in the fBackup folder (`C:\FreeFileSync\fBackup\`) using a simple name, e.g. `CopyToNas.ffs_batch`  
   While saving, check:  
   
   - `Progress window: Close automatically`  
   - `[v] Ignore errors`

3. Execute `run_fbackup.cmd` and enter the same task name.  
   This will run the configuration wizard where you can specify email settings.  
   
   - `SendEmail` → when to send reports (`Everytime` / `Never` / `OnlyError`)
   - `SmtpServer` → your SMTP server (e.g. `smtp.example.com`)
   - `Port` → SMTP port (default `465`)
   - `UseSSL` → whether to use SSL (`true` / `false`)
   - `User` → your SMTP username
   - `Password` → your SMTP password (will be encrypted)
   - `From` → sender email address
   - `To` → recipient email address
   - `SubjectPrefix` → prefix for the email subject
   
   Configuration is saved per task in an XML file called `TaskName.xml`.
   
   *Nerd note*: if SSL on port 465 fails, the script will retry with TLS on port 587.  
   You can disable this behavior by editing the script.

4. Create a new Scheduled Task with Windows Task Scheduler and select `Run a program`:
   
   - Program/script: `%windir%\system32\WindowsPowerShell\v1.0\powershell.exe`  
   - Add arguments: `-NoProfile -ExecutionPolicy Bypass -File "fbackup.ps1" -TaskName CopyToNas`  
   - Start in: `C:\FreeFileSync\fBackup`
   
   *Remember to change `CopyToNas` with the name of the task you want to run, and `C:\FreeFileSync\fBackup` with the actual path.*
