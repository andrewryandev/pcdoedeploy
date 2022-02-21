### Written by Andrew Ryan for use within NSW DoE Papercut schools ###


[cmdletbinding()]
Param()

### Script Variables ####

$PrintServer = PRINTSERVERVALUE
$LogFilePath = LOGFILEPATHVALUE
$LogFileConfirm = LOGFILECONFIRMVALUE
$RunType = $MyInvocation.MyCommand.Name

### Functions ###


function Create-Cache {

$Path = 'C:\Cache'

# test if 'C:\Cache' exists & create the directory if not
if (-not (Test-Path -Path 'C:\Cache')) {
New-Item -Path $Path -ItemType Directory
}

# add the new permissions
$acl = Get-Acl -Path $path
$permission = 'Everyone', 'FullControl', 'ContainerInherit, ObjectInherit', 'None', 'Allow'
$rule = New-Object -TypeName System.Security.AccessControl.FileSystemAccessRule -ArgumentList $permission
$acl.SetAccessRule($rule)

# set new permissions
$acl | Set-Acl -Path $path 

}


function Papercut-Startup {

$TargetFile = "\\$PrintServer\PCClient\win\pc-client-local-cache.exe"
$ShortcutFile = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\StartUp\Papercut.lnk"
$WScriptShell = New-Object -ComObject WScript.Shell
$Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
$Shortcut.TargetPath = $TargetFile
$Shortcut.Save()

}


function Papercut-GPO {

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Unregister-PSRepository -Name PSGallery
Install-PackageProvider NuGet -Force
Register-PSRepository -Default
Set-PSRepository PSGallery -InstallationPolicy Trusted
Install-Module -Name PolicyFileEditor -Confirm:$False -Force
Import-Module -Name PolicyFileEditor

$MachineDir = "$env:windir\system32\GroupPolicy\Machine\registry.pol"

$RegPath = 'Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings'
$RegName = 'ListBox_Support_ZoneMapKey'
$RegData = '1'
$RegType = 'DWord'

$RegPath2 = 'Software\Policies\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMapKey'
$RegName2 = "$PrintServer"
$RegData2 = '1'
$RegType2 = 'String'

Set-PolicyFileEntry -Path $MachineDir -Key $RegPath -ValueName $RegName -Data $RegData -Type $RegType
Set-PolicyFileEntry -Path $MachineDir -Key $RegPath2 -ValueName $RegName2 -Data $RegData2 -Type $RegType2

}


function Papercut-Trust {

$RegPath3 = "HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\Printers\PackagePointAndPrint\ListofServers"
New-Itemproperty -path $RegPath3 -Name $PrintServer -PropertyType String -Value $PrintServer -ErrorAction SilentlyContinue

}


function GP-Update {

Invoke-Command -ComputerName "$env:computername" {
$cmd1 = "cmd.exe"
$arg1 = "/c"
$arg2 = "echo y | gpupdate /force /wait:0"
&$cmd1 $arg1 $arg2
}

}


function PCDOE-Execute {

if ($LogFileConfirm -eq 'y') {

Write-Host "Setting log path..."
if (-not (Test-Path -Path "$LogFilePath")) {
New-Item -Path $LogFilePath -ItemType Directory | Out-Null
}
Write-Host "Log path set. `n"

cd $LogFilePath

Write-Host "Creating Papercut cache in C:..."
Create-Cache *> Cache.log
Write-Host "Done. `n"

Write-Host "Placing Papercut Local Cache shortcut in the common startup folder..."
Papercut-Startup *> Startup.log
Write-Host "Done. `n"

Write-Host "Setting Print Server FQDN in IE Trusted Zones..."
Papercut-GPO *> GPO.log
Write-Host "Done. `n"

Write-Host "Setting Package Point and Print Value..."
Papercut-Trust *> PackagePointPrint.log
Write-Host "Done. `n"

Write-Host "Running GPUpdate..."
GP-Update *> GPUpdate.log
Write-Host "Done. `n"
}

else {

Write-Host "Creating Papercut cache in C:..."
Create-Cache > $null
Write-Host "Done. `n"

Write-Host "Placing Papercut Local Cache shortcut in the common startup folder..."
Papercut-Startup > $null
Write-Host "Done. `n"

Write-Host "Setting Print Server FQDN in IE Trusted Zones..."
Papercut-GPO > $null
Write-Host "Done. `n"

Write-Host "Setting Package Point and Print Value..."
Papercut-Trust > $null
Write-Host "Done. `n"

Write-Host "Running GPUpdate..."
GP-Update > $null
Write-Host "Done. `n"
}

}


function PCDOE-Finish { 

Write-Host 'All Finished!'
pause

}


function Test-Verbose {
[CmdletBinding()]
param()
	[bool](Write-Verbose ([String]::Empty) 4>&1)
}


### Execute the processess ###

if ($RunType -like "*.ps1") {

$Verbosity = Test-Verbose

if ($Verbosity -eq $false) {
PCDOE-Execute
exit
}

else {
PCDOE-Execute
PCDOE-Finish
}

}

else {

Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell" -Name "ExecutionPolicy" -Value "Bypass"
Invoke-Command -Command {PCDOE-Execute; PCDOE-Finish}
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PowerShell\1\ShellIds\Microsoft.PowerShell" -Name "ExecutionPolicy" -Value "Restricted"

}

