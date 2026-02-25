# Set user impersonation protection in bulk

Impersonation protection allows you mark 350 VIP users to have them additionally protected from attacks who try to impersonate them. You can add them individually to your policies. 
![image](https://github.com/LouisMastelinck/set--TargetedUsersToProtect-bulk-script/assets/17981130/26fd00ae-dac3-471d-a1d3-b590f1045aaa)
But it contains a painful process of having to individually click all the users you want to add...
![image](https://github.com/LouisMastelinck/set--TargetedUsersToProtect-bulk-script/assets/17981130/43b359e2-21cd-41a9-be34-85b6ad47b7fc)

# **Requirement**

Permissions to manage Anti-Phish Policies

Powershell 7 --> [Windows install](https://learn.microsoft.com/en-us/powershell/scripting/install/install-powershell-on-windows?view=powershell-7.5)

[Installation of the Exchange Online PowerShell module](https://learn.microsoft.com/en-us/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-and-maintain-the-exchange-online-powershell-module)

## CSV 

This script will allow you to import a CSV that contains all of your users that you want to protect from impersonation. 
A sample CSV is provided so you know how to format it.

## Features
- Built-in error handling for Targets and enablement of the policy feature.
- Importing from a .csv file will create a "-import-results.csv" file in the same directory as the csv import file. This file will showcase all results for every entity.
- Status report in-line and in "import-results.csv" file.
- Status options:
  - Success
  - Failure
  - Already Exists (this showcases that the user was already assigned to Impersonation Protection in the selected Anti-Phish Policy

### Example of Menu
<img width="518" height="276" alt="image" src="https://github.com/user-attachments/assets/6467e06c-5b86-4478-9f35-43bf06261323" />
