# AP BILLING TOOLS

These are tools that were created to be used by allowance audit and allowance billing teams.
### Authors 
1.      Robert Todar
2.      Nicholas Ackerman

# TABLE OF CONTENTS 
   1. OVERVIEW
   2. CONFIG FILES
   3. FILE LOCATIONS 
   4. Version Control and Distribution

## Note


## Overview
 - Most all of the tools are written in VBA and are located in the file `RT_Macros.xlsb`.
 - Each user get's a copy of this workbook saved to `C:\rtmacros` on their local machine.
 - This workbook is hidden by default so it is open in the background after the user runs `RT_Macros_Box` from the command ribbon.

## Config Files

## Important File Locations 

## Version Control and Distribution
These are the instructions to save updates made to the code and optionally push out the update to users.

### How it works?
Every time a user opens up the main form in RT_Macros it call a global function `CheckForUpdates`. This function reads a config file and get's the current version that users should be on and compares it to the version they are currently using. If they do not match then the updated version is then copied over from the shared folder into the users own folder (`C:\rtmacros`).

### How to Save updates?

In the the VBE (visual basic editor) immediate window, type in `Version` and provide what type of update it is `major|minor|patch`. This will update the variable `VersionNumber` and increment it based on [semantic versioning](https://semver.org/) and save that version to the [shared folder](versions) using the format `x.x.x RT_Macros.xlsb`.
 1. In the immeidate window type in `Version` and if the update is a `Major`, `Minor`, or a `Patch`
![Immediate window](images/immediate-version.JPG)
 2. Type in `SaveUpdatedVersion` into the immediate Window. Place your Cursor within the Private Sub `Private Sub SaveUpdatedVersion()` under the `Version Control Module`. Then Press `F5`
 3. First a box will ask for the `Change Notes`. Type in a short summary and press `OK`. The box below will then appear where you can decide if you want to push out the updates to `ALL Users`, `Testers & Insiders`, or `Insiders Only`. 
![Version Control Selection](images/VersionControlSelect.JPG)

`NOTE:` - By default it will then open a prompt to see if you want to push updates to any user group. Insiders are usually just the developers while Testers are defined in the `config.json` file

### Fallback Update
If you need to fallback to a past update. Open the `config.json` file and move the version number back to a stable version. `[WARNING: This will overwrite any changes you have made since if you are not careful]` 
![Config Versions](images/configFileFallback.JPG)

### Adding New Users 
- Tool on the Misc Screen For adding More Users. Also Reference how the roles can be updated if needed.
![Add New Users](images/AddNewUser.JPG)


- ![TESTING IMage](https://github.com/Nackerman97/VBA_App_Framework/blob/4de1dc35336f670d0ff4a78e2dce5c38221726d2/images/AP_BILLING_SCREEN.jpg)