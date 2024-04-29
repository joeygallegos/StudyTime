# Read configuration from JSON file
$config = Get-Content -Raw -Path "config.json" | ConvertFrom-Json

# Outlook integration
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook" | Out-Null
$outlook = New-Object -ComObject Outlook.Application
$calendar = $outlook.Session.GetDefaultFolder(9)

# Get the current week's Monday
$today = Get-Date
$monday = $today.AddDays(-($today.DayOfWeek.value__ - [int][System.DayOfWeek]::Monday))

# Get the start and end dates of the current week (Monday to Friday)
$startOfWeek = $monday.Date
$endOfWeek = $monday.AddDays(4).AddHours(23).AddMinutes(59).AddSeconds(59).Date

# Initialize an empty array to store appointments
$appointments = @()

# Get all items in the calendar folder
$items = $calendar.Items
$meetSubject = Read-Host "What should the meetings be called?"
$meetInterval = Read-Host "30 or 60 minutes?"
$attendeesInput = Read-Host "Enter the email addresses of the attendees (comma-separated):"

# Check if the input is empty
if ([string]::IsNullOrWhiteSpace($attendeesInput)) {
    # If input is empty, initialize an empty array
    $attendees = @()
} else {
    # Split the input string by comma and trim any whitespace
    $attendees = $attendeesInput -split ',' | ForEach-Object { $_.Trim() }
}

# Iterate through each recurring meetings in the weekdays of the current week, using all calendar data
Write-Output "Task: Searching for recurring meetings to add to array.."
for ($date = $startOfWeek; $date -le $endOfWeek; $date = $date.AddDays(1)) {
    Write-Output "Appointments for $($date.ToString("dddd, MMMM dd, yyyy")):"
    
    # Iterate through items in the calendar folder
    foreach ($item in $calendar.Items | Where-Object { $_.IsRecurring }) {
        # Check if the item is an appointment and is recurring
        if ($item -is [Microsoft.Office.Interop.Outlook.AppointmentItem]) {
            # Get the recurrence pattern
            $rp = $item.GetRecurrencePattern()

            # Construct the DateTime value for the current date with the same time as the appointment start time
            $dt2 = $date
            $dt2 = $dt2.AddHours($item.Start.Hour)
            $dt2 = $dt2.AddMinutes($item.Start.Minute)
            $dt2 = $dt2.AddSeconds($item.Start.Second)

            # Attempt to retrieve the occurrence of the appointment on the current date
            try {
                # Uncomment the below line to print all attributes of the object
                $occourence = $rp.GetOccurrence($dt2)

                Write-Output ("- $($occourence.Subject) at $($occourence.Start.ToString("hh:mm tt")) (duration: $(($occourence.End - $occourence.Start).TotalMinutes))")

                # Add to array
                $appointments += [PSCustomObject]@{
                    item = $occourence
                }
            } catch [System.Runtime.InteropServices.COMException] {
                # Do nothing, no instance of the appointment falls on this date
            } catch {
                Write-Output "Error with GetOccurrence: $_"
            }
        }
    }
    
    Write-Output ""
}

Write-Output "Task: Searching for non-recurring meetings to add to array.."

# Filter items to include only appointments within the current week
$items = $items | Where-Object { $_.Start -ge $startOfWeek -and $_.End -le $endOfWeek }

# Check if the item is an appointment and is recurring
foreach ($item in $items) {
    if ($item -is [Microsoft.Office.Interop.Outlook.AppointmentItem]) {
        if ($item.IsRecurring -eq $false) {
            Write-Output ("- $($item.Subject) at $($item.Start.ToString("hh:mm tt")) (duration: $(($item.End - $item.Start).TotalMinutes))")

            # Add the AppointmentItem object to array
            $appointments += [PSCustomObject]@{
                item = $item
            }
        }
    }
}

Write-Output "Task: Visually grouping stored dates.."

# Group appointments by day
$groupedAppointments = $appointments | Group-Object { $_.item.Start.Date } | Sort-Object -Property Name

# Loop through each group (each day)
foreach ($group in $groupedAppointments) {
    Write-Output "Appointments for $($group.Name):"
    
    # Loop through appointments within the current day group
    foreach ($appointment in $group.Group) {
        Write-Output ("- $($appointment.item.Subject) from $($appointment.item.Start.ToString("hh:mm tt")) to $($appointment.item.End.ToString("hh:mm tt")) (duration: $(($appointment.item.End - $appointment.item.Start).TotalMinutes))")
    }
    
    Write-Output ""
}

# Function to create a new appointment
function CreateNewAppointment {
    param (
        [datetime]$startTime,
        [datetime]$endTime,
        [string]$subject,
        [string[]]$attendees = @() # Default value is an empty array
    )

    $outlook = New-Object -ComObject Outlook.Application
    $newAppointment = $outlook.CreateItem(1) # 1 is olAppointmentItem

    $newAppointment.Subject = $subject
    $newAppointment.Start = $startTime
    $newAppointment.End = $endTime
    $newAppointment.ReminderSet = $false

    Write-Host $startTime
    Write-Host $endTime

    # Adding attendees to the appointment
    if ($attendees.Count -ge 1) {
        foreach ($attendee in $attendees) {
            $recipient = $newAppointment.Recipients.Add($attendee)
            $recipient.Type = 1 # 1 for RequiredAttendee, 2 for OptionalAttendee
            $recipient.Resolve()
            if (-not $recipient.Resolved) {
                Write-Host "Failed to resolve attendee: $attendee"
                $recipient.Delete()
            }
        }
    }

    $newAppointment.Save()
}

# Function to check if there's a free slot for the new appointment
function FindFreeSlot {
    param (
        [array]$appointments,
        [datetime]$startDate,
        [datetime]$endDate
    )

    # Function to find available time slots
    function Get-AvailableTimeSlots {
        param (
            [datetime]$start,
            [datetime]$end,
            [int]$interval
        )
    
        $availableSlots = @()
        $currentTime = $start
    
        while ($currentTime -le $end) {
            $slotEnd = $currentTime.AddMinutes($interval)
            if ($slotEnd -le $end) {
                # Check if the slot duration matches 30 or 60 minutes
                if (($interval -eq 30 -or $interval -eq 60) -and ($currentTime.Minute % 15 -eq 0)) {
                    $availableSlots += [PSCustomObject]@{
                        Start = $currentTime
                        End = $slotEnd
                    }
                }
            }

            # Scan every 15 minutes
            $currentTime = $currentTime.AddMinutes(15) # Change interval to 15 minutes
        }
    
        return $availableSlots
    }    

    # Initialize variables to track available slots
    $interval60 = 60
    $interval30 = 30
    $availableSlots = @()

    $currentInterval = [int]$meetInterval

    if ($currentInterval -eq $interval60) {
        # Find all available 60-minute slots first
        Write-Output "Finding all available 60-minute slots.."
        $availableSlots += Get-AvailableTimeSlots -start $startDate -end $endDate -interval $currentInterval

        # If no 60-minute slots available, find 30-minute slots
        if ($availableSlots.Count -eq 0) {
            $currentInterval = $interval30
            Write-Output "Debug: No 60-minute slots available, searching for 30-minute slots.."
            $availableSlots += Get-AvailableTimeSlots -start $startDate -end $endDate -interval $currentInterval
        }
    }
    elseif ($currentInterval -eq $interval30) {
        Write-Output "Finding all available 30-minute slots.."
        $availableSlots += Get-AvailableTimeSlots -start $startDate -end $endDate -interval $currentInterval
    }
    
    # Sort available slots by start time
    $availableSlots = $availableSlots | Sort-Object -Property Start

    # TODO: Preference of padding to invites - 15 min after meeting, 15 before next - if availible
    if ($currentInterval -eq $interval60) {
        # If 60 min still in question, search for time with padding
        # 30 min with padding is not enough time

        # for each availible slot
        # find one where we can add 15 min before+after (shift up)
        # $currentTime.AddMinutes($interval)
        Write-Output "Debug: Still hunting 60 min slots"
    }
    
    # Check for conflicts with existing appointments
    foreach ($slot in $availableSlots) {
        $conflicts = $appointments | Where-Object {
            $_.item.Start -lt $slot.End -and $_.item.End -gt $slot.Start
        }

        # Write-Host "Slot - $($slot.Start) $($slot.End)"

        # No conflicts found at this timebox
        if ($conflicts.Count -eq 0) {
            Write-Output "Creating appointment from $($slot.Start.ToString()) to $($slot.End.ToString()) ($($currentInterval) min)"
            if ($attendees.Count -eq 0) {
                Write-Host "- No attendees specified!"
                # Call the function to create the appointment without attendees
                CreateNewAppointment -startTime $slot.Start -endTime $slot.End -subject $meetSubject
            } else {
                Write-Host "- Attendees specified!"
                # Call the function to create the appointment with attendees
                CreateNewAppointment -startTime $slot.Start -endTime $slot.End -subject $meetSubject -attendees $attendees
            }
            return
        }
        else {
            Write-Host "- Conflicts found, skipping"
        }
    }

    Write-Output "No available slot found on $($date.ToString("dddd, MMMM dd, yyyy"))"
}

Write-Output "Task: Hunting for free time.."

$today = Get-Date
$monday = $today.AddDays(-($today.DayOfWeek.value__ - 1))
$friday = $monday.AddDays(4)

# Retrieve start and end time from configuration
$startTime = $config.working_hours.start_time
$endTime = $config.working_hours.end_time

# Loop through days
for ($date = $monday; $date -le $friday; $date = $date.AddDays(1)) {
    $startOfDay = $date.Date.Add([TimeSpan]::Parse($startTime))
    $endOfDay = $date.Date.Add([TimeSpan]::Parse($endTime))

    FindFreeSlot -appointments $appointments -startDate $startOfDay -endDate $endOfDay
}