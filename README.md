# WsusModule
A WSUS PowerShell Module.

### Installation:
Create a folder called "Wsus" in %windir%\System32\WindowsPowerShell\v1.0\Modules or run the following PowerShell (as Admin):

`New-Item -ItemType Directory ${env:WinDir}\System32\WindowsPowerShell\v1.0\Modules\Wsus`

Download and copy wsus.psm1 to the folder you just created or run the following PowerShell (as Admin):

`Start-BitsTransfer -Source "https://raw.githubusercontent.com/jessiehowell/WsusModule/master/Wsus.psm1" -Destination "${env:WinDir}\System32\WindowsPowerShell\v1.0\Modules\Wsus"`

Run PowerShell as Admin and Import the module.

`Import-Module Wsus`

You now have two cmdlets you can use: Get-WsusUpdates and Install-WsusUpdates.


Get-WsusUpdates will list all updates available to you.


Install-WsusUpdates takes a few different options:
```
-include <KBNum> 
#Pass a comma separated list of KB numbers to install. Don't use any quotes when passing a list. Any KB Numbers not explicitly passed will not be installed. 

-exclude <KBNum> 
#Pass a comma separated list of KB numbers to NOT install. Any KB Numbers not explicitly passed will be installed.

-DownloadOnly 
#Using this option will only download available updates but NOT install them.
```
