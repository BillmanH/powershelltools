# Configuration variables
$meetingSubject = "Accenture Internal Meeting"
$clientMeetingSubject = "Client Meeting"

# "in" takes the items from your calendar and puts them in an csv.
$direction = "in"
# "out" takes the excel sheet and puts items in your calendar.
# $direction = "out"

$today = Get-Date
$BeginDelta = $today.DayOfWeek.value__ - 0
$closestSunday = $today.AddDays(-$BeginDelta).Date

$endDelta = $today.DayOfWeek.value__ - 6
$endingSunday = $today.AddDays(-$endDelta).Date


if ($direction -eq "in"){
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
    $calendar = Get-OutlookCalendar | where-object { ($_.start -gt [datetime]$closestSunday) -and ($_.start -lt [datetime]$endingSunday)}

    # Swap out the subject line to preserve confidentiality
    foreach ($j in $calendar){
        $j.subject = $meetingSubject
        $j.Location = "Microsoft Teams Meeting"
    }

    # Export that calendar to CSV.
    $calendar | Export-Csv -path ".\data\MyCalendar.csv"
}

if ($direction -eq "out"){
    # Later, get that CSV and load it into memory.
    $newCalendar = Import-CSV -Path ".\data\MyCalendar.csv"


    # Then you creat a meeting item for each item in that list
    foreach ($meet in $newCalendar){
        $ol = New-Object -ComObject Outlook.Application
        $meeting = $ol.CreateItem('olAppointmentItem')
        $meeting.Subject = $meet.subject
        $meeting.Location = $meet.location
        $meeting.MeetingStatus = [Microsoft.Office.Interop.Outlook.OlMeetingStatus]::olMeeting
        $meeting.Start = $meet.start
        $meeting.Duration = $meet.duration
        $meeting.Send()
    }
}
