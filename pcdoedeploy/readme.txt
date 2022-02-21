________________________________________________________________________________________________________________________________________________________________________________
                                              
                                                                                     
                                                                                      Papercut DoE Deploy 
                                                                                                                                                          Date Created: 9/9/2020
                                                                                                                                                         Last Modified: 12/05/2021                 
________________________________________________________________________________________________________________________________________________________________________________


Overview
---------

Papercut Doe Deploy is a tool designed for NSW Department of Education schools to easily create Papercut Local Cache deployment packages for end-user Windows devices.
The user runs the start.exe program, enters details relevant to their Papercut installation and from that generates .exe and .ps1 deployment files to run on end-user devices.
The deployment files create a shortcut to the Papercut server's "pc-client-local-cache.exe" program and places it in the device's common startup folder to run on user sign-in.
The package also sets the required GPOs and registry values so that the Papercut client software and print drivers run completely silently in the background.
The generated .exe file is designed to run via double-click on the end-user device. The .ps1 file is designed for remote installation and can be run silently.
Please see the subheadings below for more context on the included files & usage.

________________________________________________________________________________________________________________________________________________________________________________


Programs
---------


[start.exe]


Description:

start.exe is the application used to build the Papercut end-user deployment files.
Interaction will occur through a visible command window.

On program execution start.exe will check for, and attempt to download, the modules it requires (requires ps2exe - https://www.powershellgallery.com/packages/ps2exe).
Users will be prompted to input details regarding their papercut server, site and the package logfile behaviour.
Once the information has been gathered the details will be assigned to a variable and written to the final deployment package files.
start.exe will then confirm that the script has finished running and waits for the user to press a key before exiting.
The generated deployment files will be called "XXXX_pcdoedeploy.ps1" & "XXXX_pcdoedeploy.exe" (XXXX designates the school code).
These files will be placed in the root "pcdoedeploy" folder.


Input Strings:

"Input your print server name including .DETNSW.WIN:" <SERVERNAME>
-Gathers the user's print server FQDN to assign to the variable $PrintServer.
-Requires value in the format of - "XXXXaaaaaaOOOOO.DETNSW.WIN" - | XXXX = School Code   aaaaaa = AMS Location   OOOOO = Server Role & Number  e.g. 1234AR1002SP001.DETNSW.WIN

"Do you want Papercut DoE Deploy to create a log file on execution? Please input y for yes or n for no:" <LOGFILECONFIRM>
-Determines whether user wants the end deployment package to create log files when run and assigns the value to the variable $LogFileConfirm
-Requires value of either "y" or "n" | y = Yes  n = No

"If yes, please enter the desired path for the script log files. Otherwise, press enter to continue.:" <LOGFILEPATH>
-Gathers the directory path where the user wants the log files to be generated on the client devices and assigns value to the variable $LogFilePath
-Can be skipped with keypress if log files aren't needed
-Requires valid path value e.g. "C:\Temp"

"What is your school code? (This is used to name your EXE file):" <SCHOOLCODE>
-Gathers the user's schools code and assigns the value to the variable $SchoolCode
-Requires 4-digit numerical value e.g "1234"



[XXXX_pcdoedeploy.ps1]


Description:

XXXX_pcdoedeploy.ps1 is the .ps1 version of the user-generated deployment files.
This file can be deployed silently using your preferred software deployment program e.g. PDQ Deploy
Alternatively, you can produce verbose output and wait for user confirmation by using the '-Verbose' switch.
Running this script executes the "PCDOE-Execute" and "PCDOE-Finish" (Verbose Mode) functions - see function details below.



[XXXX_pcdoedeploy.exe]


Description:

XXXX_pcdoedeploy.exe runs the same deployment package as XXXX_pcdoedeploy.ps1 but in verbose mode.
It will open a command window, generate output and wait for the user to press a key when finished before exiting.
The package can be run by double-clickng the executable.

_______________________________________________________________________________________________________________________________________________________________________________


Functions
----------


[Create-Cache]

-Checks to see if the "C:\Cache" directory exists and creates it if not found
-Applies permissions to the "C:\Cache" directory to allow Papercut server to read and write to it.



[Papercut-Startup]

-Creates a shortcut to the "Papercut Local Cache" executable in the device's common startup folder
-The path to the executable looks like "\\$PrintServer\PCClient\win\pc-client-local-cache.exe"



[Papercut-GPO]

-Downloads and imports the "PolicyFileEditor" module (see link for more details - https://www.powershellgallery.com/packages/PolicyFileEditor)
-Edits and enables the "Site to Zone Assignment List" GPO to add the specified print server to the "Intranet Zone" (This is to avoid open file warnings on the client device)



[Papercut-Trust]

-Adds the print server FQDN to the Package Point and Print registry key (This is to establish the site print server as a trusted print provider)



[GP-Update]

-Runs gpupdate and applies the new machine policies



[PCDOE-Execute]

-Checks for log file settings and creates the directory if specified
-Executes the following functions: Create-Cache, Papercut-Startup, Papercut-GPO, Papercut-Trust, GP-Update



[PCDOE-Finish]

-Displays completion message and pauses the script



[Test-Verbose]

-Checks whether the script is running with the '-Verbose' switch
