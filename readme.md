# Calendar Import/Exporter

A modification of the [very brilliant code by ScriptingGuy1](https://devblogs.microsoft.com/scripting/use-powershell-to-export-outlook-calendar-information/)

## Problem:

You work for clients in multiple outlook tenants. You need to manage conflicts across multiple outlook calendars, even though you can't autoforward and you can't expect everyone to use your personal calendar. People in different tenants can't see your schedule so they don't know if you are free. 

This requires a lot of manually updating calendars, which creates error as they are across different time zones. This script automates that process. 

![Alt text](/docs/updated_meeting.png?raw=true)

This will give you the ability to block time in other calendars, without forwarding sensitive data to other tenants. The only information passed is the time so you can keep your schedules synched.
![Alt text](/docs/outlookmeeting.png?raw=true)



## Use:

The steps are outlined in detail in `syncCalendars.ps1`. You should run them line-by-line.


### 1. Export your calendar to .csv
This takes some time as it grabs the _whole calendar_, but once you have it you can filter based on the date that you need. 
```
$calendar = Get-OutlookCalendar | where-object { $_.start -gt [datetime]"6/01/2022"}
```
The `ps1` file has steps to clean out sensitive data from your meetings, replacing the subject and body with your choosing. 

Then you can save that calendar to `.csv` that you can email to yourself on the other tenant. 

```
$calendar = Export-Csv -path MyCalendar.csv 
```
### 2. Email the csv to your email at the client account.
The email isn't sensitive. It's just a list of your availability. 

### 3. run the second half of the script
```
$newCalendar = Import-CSV -Path ".\data\MyCalendar.csv"
```

Then you can run the script section that will create a calendar invite for each meeting.