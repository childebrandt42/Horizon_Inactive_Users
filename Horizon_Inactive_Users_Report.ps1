<#
.SYNOPSIS
    Horizon Inactive User Report from Horizon Events DB
.DESCRIPTION
    Horizon Inactive User Report Connects to Horizon Connection Server, Pulls Event DB informations, 
    and Entitlement lists, and then connects to Events DB, and parses out the last time a user logged on data.
    It then compares the last logon time to the current date, and if the last logon time is older than the date speicfied
    And Kicks out report via CSV or Excel.
    
    Requires SQL Server Module and ImportExcel Module

.NOTES
    Version:          1.0.0
    Author:           Chris Hildebrandt
    GitHub:           https://github.com/childebrandt42
    Twitter:          @childebrandt42
    Blog:             https://childebrandt42.blog/
    Date Created:     1/8/2024
    Date Updated:     1/15/2024
#>

#---------------------------------------------------------------------------------------------#
#                                  Script Varribles                                           #
#---------------------------------------------------------------------------------------------#

# SQL Account info
$SQLCreds = Get-Credential -Message "Enter Local SQL Account"

# Varible for how many days to look back from SQL
$SQLQueryDays = 365

# Days for last logon
$LastLogonDays = 90

# How would you like the report
$ReportType = "Excel" # CSV or Excel

# Horizon Connection Server FQDN
$HRZServerNames = @('Connection Server FQDN') # If you have multiple connection servers, add them here in format 'Connection Server FQDN','Connection Server FQDN'

# Horizon Credentials
$HRZCreds = Get-Credential -Message "Horizon Admin Account info"

# Report Name
$ReportName = "HorizonInactiveUserReport" # Report Name

# Report Save Location
$ReportSaveLocation = "C:\Reports\Usage"

#---------------------------------------------------------------------------------------------#
#                                  Powershell Modeuls                                         #
#---------------------------------------------------------------------------------------------#

# If SQL Server Module is not installed, install it
if(-not (Get-Module sqlserver -ListAvailable)){
    Write-Host "SQL Server Module is not installed, installing now"
    Install-Module sqlserver -Scope CurrentUser -Force
}

If($ReportType -eq "Excel"){
    Write-Host "Report Type is Excel"
    # If ImportExcel Module is not installed, install it
    if(-not (Get-Module ImportExcel -ListAvailable)){
        Write-Host "ImportExcel Module is not installed, installing now"
        Install-Module ImportExcel -Scope CurrentUser -Force
    }
}

# Import Modules
Import-Module ImportExcel
Import-Module sqlserver

#---------------------------------------------------------------------------------------------#
#                                  Script Logic                                               #
#---------------------------------------------------------------------------------------------#

function Get-HRHeader(){
    param($accessToken)
    return @{
        'Authorization' = 'Bearer ' + $($accessToken.access_token)
        'Content-Type' = "application/json"
    }
}

# Days to look back
$TimeBack = get-date -date $(get-date).adddays(-$SQLQueryDays) -format "yyyy-MM-dd HH:mm:ss"

# Set Date for Report
$ReportDate = Get-Date -Format MM-dd-yyy-HH-mm-ss

# Build Credential Object
$Credentials = New-Object psobject -Property @{
    username = $HRZCreds.UserName.Split('\')[1]
    password = $HRZCreds.GetNetworkCredential().Password
    domain = $HRZCreds.UserName.Split('\')[0]
}

# If Report Save Location does not exist, create it
if (-not (Test-Path $ReportSaveLocation)) {
    Write-Host "Creating Report Save Location"
    New-Item -Path $ReportSaveLocation -ItemType Directory
}

# Create Blank Arrays
$EventDatabasesInfo = @('')
$UsersorGroupsGlobalInfo = @('')
$UsersorGroupsLocalInfo = @('')

# Collect Data for each Horizon Connection Server
ForEach($HRZServer in $HRZServerNames){
    # Build URL
    $URL = "https://$HRZServer"

    # Get Access Token
    $accessToken = invoke-restmethod -Method Post -uri "$URL/rest/login" -ContentType "application/json" -Body ($Credentials | ConvertTo-Json)
    
    # Event Database Info
    Write-Host "Getting Event Database Info" -ForegroundColor Green
    $EventDataBase = Invoke-RestMethod -Method Get -uri "$url/rest/config/v1/event-database" -ContentType "application/json" -Headers (Get-HRHeader -accessToken $accessToken)
    # Event Database Info Gathering
    $EventDatabasesInfo += $EventDataBase

    # Global Users and Groups
    Write-Host "Getting Users and Groups" -ForegroundColor Green
    $UsersorGroupsGlobal = Invoke-RestMethod -Method Get -uri "$url/rest/config/v1/users-or-groups-global-summary" -ContentType "application/json" -Headers (Get-HRHeader -accessToken $accessToken)
    # Gather User Information for all
    $UsersorGroupsGlobalInfo += $UsersorGroupsGlobal

    # Local Users and Groups
    Write-Host "Getting local Users and Groups" -ForegroundColor Green
    $UsersorGroupsLocal = Invoke-RestMethod -Method Get -uri "$url/rest/config/v1/users-or-groups-local-summary" -ContentType "application/json" -Headers (Get-HRHeader -accessToken $accessToken)
    # Gather User Information for all
    $UsersorGroupsLocalInfo += $UsersorGroupsLocal
}


# Create Blank Arrays
$EventsHST = @('')
$SQLServer = @('')
$SQLQueryHST = @('')
$EventDataHST = @('')

# Do SQL Query for each Event Database
foreach($SQLServer in $EventDatabasesInfo){
    If($SQLServer){
        Write-Host "SQL Server Name - $($SQLServer.server_name)"
        # Define SQL Query
        $SQLQueryHST = "SELECT * from $($SQLServer.database_name).event_historical where (EventType = 'AGENT_CONNECTED') and (Time > '$TimeBack') order by time desc"

        Write-Host "Processing $($SQLServer.server_name) for Historical Events" -ForegroundColor Green
        $EventsHST += Invoke-Sqlcmd -Credential $SQLCreds -ServerInstance $($SQLServer.server_name) -Database $($SQLServer.database_name) -Query $SQLQueryHST -TrustServerCertificate | Select-Object ModuleAndEventText, Time, Node, DesktopId
    }
}

# Create Historical Event Data Array
Foreach ($EventHST in $EventsHST){
    if($EventHST){
        $UsernameHSTEDT = ''
        $UsernameHSTEDT = $EventHST.ModuleAndEventText | Out-String
        $UsernameHSTEDT = $UsernameHSTEDT.Trim('User ')
        $UsernameHSTEDT = $UsernameHSTEDT.substring(0,$UsernameHSTEDT.IndexOf(' '))
        $UsernameHSTEDT = $UsernameHSTEDT.Split('\')[$($UsernameHSTEDT.Split('\').Count-1)]
        $UsernameHSTEDT = $UsernameHSTEDT.Split('\')[$($UsernameHSTEDT.Split('\').Count-1)]

        # Create Historical Event Data Array to Export
        if ($EventHST){
            if($UsernameHSTEDT -notin $EventDataHST.UserName){
                Write-Host "Adding to user $UsernameHSTEDT to Table" -ForegroundColor Green
                $DateTime = [DateTime]::Parse($EventHST.Time)
                
                $EventDataHST += [pscustomobject]@{
                    UserName = $UsernameHSTEDT
                    LogonTime = $DateTime
                    NodeID = $EventHST.Node
                    #DesktopID = $EventHST.DesktopId
                }
            }
        }
    }

}

# Create Blank Arrays
$AllUsers = @('') 
$UserOrGroup = @('')
$AllUsersandGroups = @('')

# Combine Arrays
$AllUsersandGroups += $UsersorGroupsGlobalInfo
$AllUsersandGroups += $UsersorGroupsLocalInfo

# Remove Duplicates
$AllUsersandGroupsUnique = $AllUsersandGroups | Select-Object * -Unique

# Loop through all users and groups
Foreach($UserOrGroup in ($AllUsersandGroupsUnique)){
    If($UserOrGroup){
        $commaSeparatedGD = @('')
        $commaSeparatedGA = @('')
        $commaSeparatedLD = @('')
        $commaSeparatedLA = @('')

        $GlobalDesktopName = @('')
        $GlobalApplicationName = @('')
        $GlobalApplicationName = @('')
        $LocalDesktopName = @('')
        $LocalApplicationName = @('')

        # Get Global Desktop Entitlements
        If($UserOrGroup.global_desktop_entitlements){
            foreach($GlobalDesktopEntitlement in $UserOrGroup.global_desktop_entitlements){
                If($GlobalDesktopEntitlement){
                    $GlobalDesktopName += Invoke-RestMethod -Method Get -uri "$url/rest/inventory/v1/global-desktop-entitlements/$GlobalDesktopEntitlement" -ContentType "application/json" -Headers (Get-HRHeader -accessToken $accessToken)
                }
            }
        }
        # Get Global Application Entitlements
        If($UserOrGroup.global_application_entitlements){
            foreach($GlobalApplicationEntitlement in $UserOrGroup.global_application_entitlements){
                If($GlobalApplicationEntitlement){
                    $GlobalApplicationName += Invoke-RestMethod -Method Get -uri "$url/rest/inventory/v1/global-application-entitlements/$GlobalApplicationEntitlement" -ContentType "application/json" -Headers (Get-HRHeader -accessToken $accessToken)
                }
            }
        }
        # Get Local Desktop Entitlements
        If($UserorGroup.desktop_pool_ids){
            foreach($LocalDesktopEntitlement in $UserOrGroup.desktop_pool_ids){
                If($LocalDesktopEntitlement){
                    $LocalDesktopName += Invoke-RestMethod -Method Get -uri "$url/rest/inventory/v1/desktop-pools/$LocalDesktopEntitlement" -ContentType "application/json" -Headers (Get-HRHeader -accessToken $accessToken)
                }
            }
        }
        # Get Local Application Entitlements
        If($UserorGroup.application_pool_ids){
            foreach($LocalApplicationEntitlement in $UserOrGroup.application_pool_ids){
                If($LocalApplicationEntitlement){
                    $LocalApplicationName += Invoke-RestMethod -Method Get -uri "$url/rest/inventory/v1/application-pools/$LocalApplicationEntitlement" -ContentType "application/json" -Headers (Get-HRHeader -accessToken $accessToken)
                }
            }
        }

        # Create Comma Separated Lists for each Global Desktop and Application Entitlements
        $commaSeparatedGD = (($GlobalDesktopName | ForEach-Object { $_.name }) -join ',').TrimStart(",")
        $commaSeparatedGA = (($GlobalApplicationName | ForEach-Object { $_.name }) -join ',').TrimStart(",")
        $commaSeparatedLD = (($LocalDesktopName | ForEach-Object { $_.name }) -join ',').TrimStart(",")
        $commaSeparatedLA = (($LocalApplicationName | ForEach-Object { $_.name }) -join ',').TrimStart(",")

        # Get Group Info
        If($UserOrGroup.Group -like 'True'){
            Write-Host "Gathering Group Info for - $($UserOrGroup.name)" -ForegroundColor Green
            $AllUsers += Get-ADGroupMember -Identity $($UserOrGroup.Name) -Server $UserOrGroup.domain -Credential $HRZCreds | Get-ADUser -Properties * | select-object Name,SamAccountName,Enabled,Mail,@{n="AD_Group"; e={$UserOrGroup.Name}},@{n="Domain Name"; e={($UserOrGroup.domain)}},@{n="AD Group Name"; e={$($UserOrGroup.Name)}},Department,Title,Country,Office,City,State,StreetAddress,@{Name="Manager Name"; Expression={(Get-ADUser -Identity $_.Manager -Server $UserOrGroup.domain -Credential $HRZCreds).Name}},@{Name="Manager Email"; Expression={(Get-ADUser -Identity $_.Manager -Properties mail -Server $UserOrGroup.domain -Credential $HRZCreds).mail}},@{n="Global Desktop Entitlements"; e={$commaSeparatedGD}},@{n="Global Application Entitlements"; e={$commaSeparatedGA}},@{n="Local Desktop Entitlements"; e={$commaSeparatedLD}},@{n="Local Application Entitlements"; e={$commaSeparatedLA}}
        }
        # Get User Info
        if($UserOrGroup.Group -like 'False'){
            Write-Host "Gathering User Info for - $($UserOrGroup.name)" -ForegroundColor Green
            $GroupName = 'Direct User Assigned'
            $AllUsers += Get-ADUser -Identity $($UserOrGroup.Name) -Server $UserOrGroup.domain -Credential $HRZCreds | select-object Name,SamAccountName,Enabled,Mail,@{n="AD_Group"; e={$GroupName}},@{n="Domain Name"; e={($UserOrGroup.domain)}},@{n="AD Group Name"; e={$($GroupName)}},Department,Title,Country,Office,City,State,StreetAddress,@{Name="Manager Name"; Expression={(Get-ADUser -Identity $_.Manager -Server $UserOrGroup.domain -Credential $HRZCreds).Name}},@{Name="Manager Email"; Expression={(Get-ADUser -Identity $_.Manager -Properties mail -Server $UserOrGroup.domain -Credential $HRZCreds).mail}},@{n="Global Desktop Entitlements"; e={$commaSeparatedGD}},@{n="Global Application Entitlements"; e={$commaSeparatedGA}},@{n="Local Desktop Entitlements"; e={$commaSeparatedLD}},@{n="Local Application Entitlements"; e={$commaSeparatedLA}}

        }
    }

}

# Remove Duplicates
$AllUsersUnique = $AllUsers | Select-Object * -Unique

# Build Past Days Varrible
$PastDays = get-date -date $(get-date).adddays(-$LastLogonDays) -format "yyyy-MM-dd HH:mm:ss"

# Create Blank Arrays
$UserLogOnDataFull = @('')
$ActiveUserLogOnDataFull = @('')
$UserLogonData = @('')

# Loop through all users and groups
Foreach($User in $AllUsersUnique){
    if($User){
        $UserLogonData = $EventDataHST | Where-Object {$_.UserName -like $User.SamAccountName}
        If($UserLogonData.LogonTime -lt $PastDays){
            Write-Host "Questionable User - $($User.Name)" -ForegroundColor Red
            If ($null -eq $UserLogonData.LogonTime){
                $LogonTime = 'Never Logged On'
            }else{
                $LogonTime = $UserLogonData.LogonTime
            }
            $UserLogOnDataFull += [pscustomobject]@{
                UserLogOnName = $UserLogonData.UserName
                LastLogonTime = $LogonTime
                NodeID = $UserLogonData.NodeID
                UserName = $User.Name
                SamAccountName = $User.SamAccountName
                Enabled = $User.Enabled
                Mail = $User.Mail
                AD_Group = $User.AD_Group
                Domain_Name = $User.'Domain Name'
                AD_Group_Name = $User.'AD Group Name'
                Department = $User.Department
                Title = $User.Title
                Country = $user.Country
                Office = $User.Office
                City = $User.City
                State = $User.State
                StreetAddress = $User.StreetAddress
                Manager_Name = $User.'Manager Name'
                Manager_Email = $user.'Manager Email'
                Global_Desktop_Entitlements = $User.'Global Desktop Entitlements'
                Global_Application_Entitlements = $User.'Global Application Entitlements'
                Local_Desktop_Entitlements = $User.'Local Desktop Entitlements'
                Local_Application_Entitlements = $User.'Local Application Entitlements'
            }
        
        }else{
            Write-Host "Productive User $($User.Name)" -ForegroundColor Green
            If ($null -eq $UserLogonData.LogonTime){
                $LogonTime = 'Never Logged On'
            }else{
                $LogonTime = $UserLogonData.LogonTime
            }
            $ActiveUserLogOnDataFull += [pscustomobject]@{
                UserLogOnName = $UserLogonData.UserName
                LastLogonTime = $LogonTime
                NodeID = $UserLogonData.NodeID
                UserName = $User.Name
                SamAccountName = $User.SamAccountName
                Enabled = $User.Enabled
                Mail = $User.Mail
                AD_Group = $User.AD_Group
                Domain_Name = $User.'Domain Name'
                AD_Group_Name = $User.'AD Group Name'
                Department = $User.Department
                Title = $User.Title
                Country = $user.Country
                Office = $User.Office
                City = $User.City
                State = $User.State
                StreetAddress = $User.StreetAddress
                Manager_Name = $User.'Manager Name'
                Manager_Email = $user.'Manager Email'
                Global_Desktop_Entitlements = $User.'Global Desktop Entitlements'
                Global_Application_Entitlements = $User.'Global Application Entitlements'
                Local_Desktop_Entitlements = $User.'Local Desktop Entitlements'
                Local_Application_Entitlements = $User.'Local Application Entitlements'
            }
        }
    }
}

# Export your Writables Data
if($ReportType -eq 'Excel'){
    $UserLogOnDataFull | Select-Object ('UserLogOnName','LastLogonTime','NodeID','UserName','SamAccountName','Enabled','Mail','AD_Group','Domain_Name','AD_Group_Name','Department','Title','Country','Office','City','State','StreetAddress','Manager_Name','Manager_Email','Global_Desktop_Entitlements','Global_Application_Entitlements','Local_Desktop_Entitlements','Local_Application_Entitlements') | Export-Excel -Path "$ReportSaveLocation\$ReportName-$ReportDate.xlsx" -WorksheetName 'Horizon Inactive Usage Report' -AutoSize
    $ActiveUserLogOnDataFull | Select-Object ('UserLogOnName','LastLogonTime','NodeID','UserName','SamAccountName','Enabled','Mail','AD_Group','Domain_Name','AD_Group_Name','Department','Title','Country','Office','City','State','StreetAddress','Manager_Name','Manager_Email','Global_Desktop_Entitlements','Global_Application_Entitlements','Local_Desktop_Entitlements','Local_Application_Entitlements') | Export-Excel -Path "$ReportSaveLocation\$ReportName-$ReportDate.xlsx" -WorksheetName 'Horizon Active Usage Report' -AutoSize -Append
}else{
    $UserLogOnDataFull | Select-Object ('UserLogOnName','LastLogonTime','NodeID','UserName','SamAccountName','Enabled','Mail','AD_Group','Domain_Name','AD_Group_Name','Department','Title','Country','Office','City','State','StreetAddress','Manager_Name','Manager_Email','Global_Desktop_Entitlements','Global_Application_Entitlements','Local_Desktop_Entitlements','Local_Application_Entitlements') | Export-Csv -Path "$ReportSaveLocation\$ReportName-$ReportDate.csv" -NoTypeInformation
}