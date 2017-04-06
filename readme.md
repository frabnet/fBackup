## fBackup

This tool is a helper for FreeFileSync. It can send you email with FreeFileSync logs and automate some backup related tasks.
I currently use it on 10-20 servers and constantly try to improve/fix bugs.

## Download

You can download fBackup directly from GitHub.
It requires FreeFileSync and PowerShell v4 or greater. Check Installation steps above.


## Installation
* Check current PowerShell version with the command `$PSVersionTable`.
    fBackup is tested on PowerShell 4 or greater. You can download v4 here: https://www.microsoft.com/en-us/download/details.aspx?id=40855
* Download and install FreeFileSync: https://www.freefilesync.org/download.php
    __Be careful while installing__: [there are advertising](https://www.freefilesync.org/faq.php#silent-ad) which I like to avoid.
    I suggest to install it in a convenient directory like `C:\FreeFileSync\`
* Create a Logs folder inside FreeFileSync folder, for example: `C:\FreeFileSync\Logs`
* Unzip fBackup inside FreeFileSync folder: `C:\FreeFileSync\fBackup`

## Task Creation

* Create a new Job with FreeFileSync and save it as Batch Job.
    While saving, set options:
     ```
    Handle errors: Ignore
    On completion: Close progress dialog
    Save log: C:\FreeFileSync\Logs\
    Limit: 100
    ```
    Save it in `C:\FreeFileSync\fBackup\` with a simple name, like `CopyToNas.ffs_batch`
    You can refer to FreeFileSync online [manual](https://www.freefilesync.org/manual.php?topic=schedule-a-batch-job) for more information.
* Copy `fBackup\example_task.ini` file with the same name as previously created task (like `CopyToNas.ini`) and edit it to match your needs.
* Test email?
* Create new Schedule Task with Windows Scheduler
    ```
    c:\windows\system32\WindowsPowerShell\v1.0\powershell.exe
    -NoProfile -Executionpolicy Bypass -File "c:\FreeFileSync\fbackup\backup.ps1" -nomeTask CopyToNas
    c:\FreeFileSync\fbackup
    ```
