## fBackup

This tool is a helper for FreeFileSync. It can send you email with FreeFileSync logs and automate some backup related tasks.
I currently use it on 10-20 servers and constantly try to improve/fix bugs.

## Download

You can download fBackup directly from GitHub.
It requires FreeFileSync and PowerShell v4 or greater. Check Installation steps above.

## Installation

* Check current PowerShell version with the command `Get-Host | Select-Object Version`.
    fBackup is tested on PowerShell 4 or greater. You can download v5 from [Microsoft website](https://aka.ms/wmf51download).
* Download and install FreeFileSync: https://www.freefilesync.org/download.php
    I suggest to install it in a convenient directory like `C:\FreeFileSync\`    
* Create a Logs folder inside FreeFileSync folder, for example: `C:\FreeFileSync\Logs`
* Unzip fBackup inside FreeFileSync folder: `C:\FreeFileSync\fBackup`

## Task Creation

* Create a new Job with FreeFileSync as you like.
* In syncronization settings, please check:
`[v] Replace default log path`
And set the folder to ``C:\FreeFileSync\Logs\``
* Save it as a Batch Job in `C:\FreeFileSync\fBackup\` with a simple name, like `CopyToNas.ffs_batch`
* While saving, please check: ``Progress window: Close automatically`` and ``[v] Ignore errors``
* Copy `fBackup\example_task.ini` and rename it the same name as the Batch Job you just saved (eg: `CopyToNas.ini`). Edit it to match your needs.
* Optional but suggested: You can test the task by running `test_task.cmd` and entering the name of the task you want to run.
    __NOTE: This will also run FreeFileSync Job!__ To test a task without executing FreeFileSync (testing only E-Mail settings), you can set Enable=Test inside the .ini file. Remember to set it back after the test with the value you want it to be.
* Create new Schedule Task with Windows Scheduler and select to Run a program:
	```
	Program/script: c:\windows\system32\WindowsPowerShell\v1.0\powershell.exe
	Add arguments: -NoProfile -Executionpolicy Bypass -File "backup.ps1" -nomeTask CopyToNas
	Start in: c:\FreeFileSync\fbackup
	```
	*Change `CopyToNas` with the name of the task you want to run.*
