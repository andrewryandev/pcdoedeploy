### Load PS2EXE Module ###

function InstallPS2EXE-Module {
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
Install-PackageProvider NuGet -Force
Set-PSRepository PSGallery -InstallationPolicy Trusted
Install-Module -Name ps2exe -Confirm:$False -Force
Import-Module -Name ps2exe
}


Write-Host "Installing Dependencies..."
If (-not(Get-InstalledModule ps2exe -ErrorAction silentlycontinue)) {
  InstallPS2EXE-Module | Out-Null
}
Else {
  Import-Module -Name ps2exe | Out-Null
}
Write-Host "Dependencies installed successfully `n"

### Collect Script Values ###

Write-Host "Collecting Script Info... `n"

$PrintServerName = Read-Host -Prompt 'Input your print server name including .DETNSW.WIN'
$LogConfirm = Read-Host -Prompt 'Do you want Papercut DoE Deploy to create a log file on execution? Please input y for yes or n for no'
$LogPath = Read-Host -Prompt 'If yes, please enter the desired path for the script log files. Otherwise, press enter to continue.'
$SchoolCode = Read-Host -Prompt 'What is your school code? (This is used to name your EXE file)'
$ProgramBase = "_pcdoedeploy"
$EXESuffix = ".exe"
$PSSuffix = ".ps1"
$EXEName = $SchoolCode+$ProgramBase+$EXESuffix
$PSName = $SchoolCode+$ProgramBase+$PSSuffix


### Apply Template Script Values ###


((Get-Content -path .\resources\pcdoedeploy_master.ps1 -Raw) -replace 'PRINTSERVERVALUE',"`"$PrintServerName`"" -Replace 'LOGFILEPATHVALUE',"`"$LogPath`"" -Replace 'LOGFILECONFIRMVALUE',"`"$LogConfirm`"") | Set-Content -Path .\resources\pcdoedeploy_defined.ps1
Write-Host "Information Collected and Applied To Template Script `n"

### Build EXE ###

Write-Host "Building PS Script and EXE File... `n"

Invoke-ps2exe ".\resources\pcdoedeploy_defined.ps1" ".\$EXEName" -requireAdmin

((Get-Content -path .\resources\pcdoedeploy_defined.ps1 -Raw) -replace 'Write-Host','Write-Verbose -Message') | Set-Content -Path .\$PSName

Write-Host "`n Your PaperCut admin EXE and PS script files have been created and are ready for deployment.`n The EXE file can be run directly on the target device and will write output to a visible Powershell session.`n The ps1 file can be deployed silently using the standard arguments. `n Alternatively you can run the ps1 file with the '-Verbose' switch to enable visible output and wait for user interaction before closing. `n"

### Wait For User To Respond ####

Write-Host -NoNewLine 'Press any key to exit...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
