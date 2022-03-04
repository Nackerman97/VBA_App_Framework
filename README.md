# AP BILLING TOOLS

These are tools that were created to be used by Allowance Audit for Accounts Payable to help with the billing of Bouncers

### Authors 
1. Robert Todar
2. Nicholas Ackerman

# TABLE OF CONTENTS 
   1. OVERVIEW
   2. CONFIG FILES
   3. Important File Locations 
   4. Version Control and Distribution
   5. AP BILLING TOOL FUNCTIONS

# Overview
 - Most all of the tools are written in VBA and are located in the file `AP_BILLING_TOOL.xlsb`.
 - Each user get's a copy of this workbook saved to `C:\AP_BILLING_TOOL` on their local machine.
 - This workbook is hidden by default so it is open in the background after the user runs `AP_Macros_Box` from the command ribbon.

## Config Files
- [AP Billing Tool Config](config.json) - This file used to modify active versions for the users
- [AP Billing Tools Users](users.json) - This file is used to show the current users for the tool and their version number
- [Bouncer Billing Tool Config](/sql/Bouncer/ConfigBouncer.json) - This file is used to make modifcations to the Bouncer Billing processes in the tool

## Important File Locations 
- [Bouncer Billing Tool SQL](sql/) - This folder location contains the sql for the tool
- [Bouncer Billing Tool Versions](versions/) - This folder location contains the backup versions for the tool
- [Bouncer Billing Tool Change Log](versions/ChangeLog.log) - The file contains a log of all of the updates pushed out the application.
- [Bouncer Billing Tool Analytics](data/analytics/) - This folder contains a list of dated txt files that outline how many times a function has been run from the AP Billing package
- [Bouncer Billing Tool Note](data/Note.txt) - This file is used to highlight something important by displaying a message at the bottom of the main screen for the tool
- [Bouncer Billing Tool Allocation Log](data/logs/AUTOMATED_ALLOCATIONS/) - This folder contains a log of all the allocations completed using the tool
- [Bouncer Billing Tool Billing Log](data/logs/AUTOMATED_BILLINGS/) - This folder contains a log of all the billings completed using the tool
- [Bouncer Billing Tool Environemnt File](.env) - This file is an environment data file and is used to temporarily hold information such as file locations or testing values.
- [Bouncer Billing Tool Images](images/) - This folder contains images that are used to operate the application as well as populate the images in this README

# Version Control and Distribution
These are the instructions to save updates made to the code and optionally push out the update to users.

## How it works?
Every time a user opens up the main form in AP_BILLING_TOOLS it call a global function `CheckForUpdates`. This function reads a config file and get's the current version that users should be on and compares it to the version they are currently using. If they do not match then the updated version is then copied over from the shared folder into the users own folder (`C:\AP_BILLING_TOOL`).

## How to Save updates?

   ![App Configuration File](https://github.com/Nackerman97/VBA_App_Framework/blob/eb9c22940df8cccf71f946dfac338d3640e0d795/images/ConfigFile.jpg)

1. If you are listed as a developer you will have an additional set of options `(Shown Below)` to push out update to the application. You can add developers by modifying the configuration file manually in the `developer` section of the ConfigFile shown above. 

   ![Version Control Selection](https://github.com/Nackerman97/VBA_App_Framework/blob/eb9c22940df8cccf71f946dfac338d3640e0d795/images/DEVELOPMENT.jpg)

2. A developer can add new users to the tool by pressing `Add User` and typing in the users `USERNAME` in ALL CAPS.
3. A developer can push out updates/modifcations made to the VBA by selecting `Push Update`. This will prompt the field below. Provide what type of update it is `major|minor|patch`. This will update the variable `VersionNumber` and increment it based on [semantic versioning](https://semver.org/) and save that version to the [shared folder](versions) using the format `x.x.x AP_BILLING_TOOL.xlsb`. `Patchs` are for small bugs, `Minor` is for larger functions added for an existing function, `Major` is for adding a new capability.

   ![Version Control Selection](https://github.com/Nackerman97/VBA_App_Framework/blob/eb9c22940df8cccf71f946dfac338d3640e0d795/images/Version_Control.jpg)

4. After you select the Version Type the program the program will prompt the field below. This application has the ability to push out update to only specific users. If you are developing an application and you only want to push out the update to testers you can select the option for `Testers & Insiders`. However, if you want to push out an update to all users you can select `Push to all users`

   ![Version Control Selection](https://github.com/Nackerman97/VBA_App_Framework/blob/eb9c22940df8cccf71f946dfac338d3640e0d795/images/VersionControlSelect.jpg)

`NOTE:` - By default it will then open a prompt to see if you want to push updates to any user group. Insiders are usually just the developers while Testers are defined in the `config.json` file


## Fallback Update
If you need to fallback to a past update. Open the `config.json` file and move the version number back to a stable version. `[WARNING: This will overwrite any changes you have made since if you are not careful]` 

# AP BILLING TOOL FUNCTIONS

Below are the functions that exist in the tools package.

## Bouncer Data Lookup
   These functions are used for pulling bouncer information 

### Functions
 1. PULL BOUNCER CASE OI
 2. PULL BOUNCER SHORTSHIPMENT
 3. PULL BOUNCER SWELL / FREIGHT
 4. PULL BOUNCER VENDOR COST

## Bouncer Billing 
   These functions are used billing bouncers

### Functions
 1. BILL BOUNCER OI
 2. BILL BOUNCER 
 3. BILL SWELL FREIGHT AUTO
