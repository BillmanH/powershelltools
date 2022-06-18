# Configuration variables
$meetingSubject = "Accenture Internal Meeting"
$clientMeetingSubject = "Client Meeting"
$meetingBody = @"
I have a meeting at this time within another 0365 tenant. I'll be offline at this time.

I'll respond as soon as I'm able to return.
"@
$startdate = "6/19/2022"
$enddate = "6/21/2022"



# Get the calendar
Function Get-OutlookCalendar
{
    Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
    $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
    $outlook = new-object -comobject outlook.application
    $namespace = $outlook.GetNameSpace("MAPI")
    $folder = $namespace.getDefaultFolder($olFolders::olFolderCalendar)
    $folder.items |
    Select-Object -Property Subject, Start, End, Duration, Location
}
$calendar = Get-OutlookCalendar | where-object { ($_.start -gt [datetime]$startdate) -and ($_.start -lt [datetime]$enddate)}

# Swap out the subject line to preserve confidentiality
foreach ($j in $calendar){
    $j.subject = $meetingSubject
    $j.Location = "Microsoft Teams Meeting"
    $j.Body = $meetingBody
}

# Export that calendar to CSV.
$calendar | Export-Csv -path ".\data\MyCalendar.csv"


# Later, get that CSV and load it into memory.
$newCalendar = Import-CSV -Path ".\data\MyCalendar.csv"


# Then you creat a meeting item for each item in that list
foreach ($meet in $newCalendar){
    $ol = New-Object -ComObject Outlook.Application
    $meeting = $ol.CreateItem('olAppointmentItem')
    $meeting.Subject = $meet.subject
    $meeting.Body = $meet.Body
    $meeting.Location = $meet.location
    $meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
    $meeting.Start = $meet.start
    $meeting.Duration = $meet.duration
    $meeting.Send()
}
