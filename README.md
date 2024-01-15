<p align="center">
    <a href="https://twitter.com/childebrandt42" alt="Twitter">
            <img src="https://img.shields.io/twitter/follow/Childebrandt42.svg?style=social"/></a>
</p>
<!-- ********** DO NOT EDIT THESE LINKS ********** -->

# VMware Horizon Inactive User Report

VMware Horizon Inactive User Report which works in conjunction with [SQLServer Powershell Module](https://www.powershellgallery.com/packages/SqlServer/22.1.1) and [ImportExcel Powershell Module](https://github.com/dfinke/ImportExcel).

Please refer to my blog [Blog Website](https://www.childebrandt42.blog) for more detailed information about this project. 

Blog post for this project [Blog Post](https://childebrandt42.blog/2024/01/15/in-depth-user-inactivity-analysis-for-vmware-horizon-environment)

# :books: Sample Reports

## Sample Report
Sample Horizon Usage Report Excel format: [Sample-HorizonUserReport.xlsx](https://htmlpreview.github.io/?https://raw.githubusercontent.com/childebrandt42/Horizon_Inactive_Users/main/Samples/Sample-HorizonUserReport.xlsx)

# :beginner: Getting Started
Below are the instructions on how to run the VMware Horizon Usage Report

### PowerShell
This report is compatible with the following PowerShell versions;

<!-- ********** Update supported PowerShell versions ********** -->
| Windows PowerShell 5.1 |     PowerShell 7    |
|:----------------------:|:--------------------:|
|   :white_check_mark:   | :white_check_mark: |
## :wrench: System Requirements
<!-- ********** Update system requirements ********** -->
PowerShell 5.1 or PowerShell 7, and the following PowerShell modules are required for generating a VMware Horizon Usage Report.

- [SQL Server Module](https://www.powershellgallery.com/packages/SqlServer/)
- [Import Excel Module](https://www.powershellgallery.com/packages/ImportExcel/)

## :package: Instructions

Download script

Place script on machine that can talk to Horizon Events DB SQL server

Fill in the Varribles:  
* $SQLCreds - Credentials for SQL Table
* $SQLQueryDays - Days to Query back in SQL Server
* $LastLogonDays - Last Logon Days for report. 
* $ReportType - Report Type, Either Excel or CSV
* $HRZServerNames - Connection Server FQDN, only need one connection server per cluster. 
* $HRZCreds - Horizon Admin Account info
* $ReportName - Report Name
* $ReportSaveLocation - Report Save Location


Then Run the script. 
