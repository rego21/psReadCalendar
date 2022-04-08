<#
.SYNOPSIS
	Opens and reads Outlook Calendar Events
.DESCRIPTION
	This PowerShell script reads all Outlook Calendar events based on a query.
    The default query fetch all the Portuguese Holidays (the Holiday events can be added in the calendar, Outlook provides them).
    The query is easly editiable.
    The program outputs all the events with the respective information (start date, subject, location, ...).
.EXAMPLE
	PS> ./readCalendar
.LINK
	https://github.com/rego21/psReadCalendar
.NOTES
	Author: Miguel Rego
#>

# Opens and returns all Outlook Calendar Events
Function Get-OutlookCalendar
{
 Add-type -assembly "Microsoft.Office.Interop.Outlook" | out-null
 $olFolders = "Microsoft.Office.Interop.Outlook.OlDefaultFolders" -as [type]
 $outlook = new-object -comobject outlook.application
 $namespace = $outlook.GetNameSpace("MAPI")
 $folder = $namespace.getDefaultFolder($olFolders::olFolderCalendar)
 $folder.items
}

# Get current year
# TODO: Add year as parameter
$year = (Get-Date).year;

# Get Portuguese Holidays.
$daysOffOrigin = Get-OutlookCalendar | where-object {
     $_.Start -gt [datetime]"1 January $year" -AND $_.Start -lt [datetime]"31 December $year" -AND
     ($_.Categories -eq "Holiday" -AND $_.Location -eq "Portugal" -AND $_.Subject -ne "Holy Thursday")
     }

# This is needed in order to do a deep copy.
$daysOff = @()
foreach ($item in $daysOffOrigin) {

        $day = $item | Select-Object Start, End, Subject, Categories | ConvertTo-Json -depth 100 | ConvertFrom-Json
        # These "ternary" operations are needed, because Holidays start and end dates give us two days...
        $daysOff += [PSCustomObject]@{
        Start = $day.start
        Subject = $day.Subject
        End =  $day.end
    }
}

# Print all the Off days [debug]
foreach ($item in $daysOff) {
    $item.start
    $item.end
    $item.Subject
}
