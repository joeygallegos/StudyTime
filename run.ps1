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

# Iterate through each recurring meetings in the weekdays of the current week
Write-Output "Task: Searching for recurring meetings to add to array.."
for ($date = $startOfWeek; $date -le $endOfWeek; $date = $date.AddDays(1)) {
    Write-Output "Appointments for $($date.ToString("dddd, MMMM dd, yyyy")):"
    
    # Iterate through items in the calendar folder
    foreach ($item in $calendar.Items) {
        # Check if the item is an appointment and is recurring
        if ($item -is [Microsoft.Office.Interop.Outlook.AppointmentItem] -and $item.IsRecurring) {
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
                # Write-Output ("- $($item.Subject) at $($dt2.ToString("hh:mm tt"))")

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

# Get all items in the calendar folder
$items = $calendar.Items

# Filter items to include only appointments within the current week
$items = $items | Where-Object { $_.Start -ge $startOfWeek -and $_.End -le $endOfWeek }

Write-Output "Task: Searching for non-recurring meetings to add to array.."

# Check if the item is an appointment and is recurring
foreach ($item in $items) {
    if ($item -is [Microsoft.Office.Interop.Outlook.AppointmentItem]) {
        if ($item.IsRecurring -eq $false) {
            Write-Output ("- $($item.Subject) at $($item.Start.ToString("hh:mm tt")) (duration: $(($item.End - $item.Start).TotalMinutes))")
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
        [datetime]$endTime
    )

    $newAppointment = $outlook.CreateItem(1) # 1 is olAppointmentItem

    $newAppointment.Subject = "Study More"
    $newAppointment.Start = $startTime
    $newAppointment.End = $endTime
    $newAppointment.ReminderSet = $false

    $newAppointment.Save()
}

# Function to check if there's a free slot for the new appointment
function FindFreeSlot {
    param (
        [array]$appointments,
        [datetime]$startDate,
        [datetime]$endDate
    )

    # Function to calculate duration between two datetimes
    function Get-TimeSpanMinutes($start, $end) {
        return ($end - $start).TotalMinutes
    }

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
                $availableSlots += [PSCustomObject]@{
                    Start = $currentTime
                    End = $slotEnd
                }
            }
            $currentTime = $currentTime.AddMinutes($interval)
        }

        return $availableSlots
    }

    # Initialize variables to track available slots
    $interval60 = 60
    $interval30 = 30
    $availableSlots = @()

    # Find all available 60-minute slots first
    $availableSlots += Get-AvailableTimeSlots -start $startDate -end $endDate -interval $interval60

    # If no 60-minute slots available, find 30-minute slots
    if ($availableSlots.Count -eq 0) {
        $availableSlots += Get-AvailableTimeSlots -start $startDate -end $endDate -interval $interval30
    }

    # Sort available slots by start time
    $availableSlots = $availableSlots | Sort-Object -Property Start

    # TODO: Add padding to invites - 15 min after meeting, 15 before next - if availible
    if ($interval -eq $interval60) {
        # If 60 min still in question, search for time with padding
        # 30 min with padding is not enough time
        Write-Output "still 60 hunt"
    }
    
    # Check for conflicts with existing appointments
    foreach ($slot in $availableSlots) {
        $conflicts = $appointments | Where-Object {
            $_.item.Start -lt $slot.End -and $_.item.End -gt $slot.Start
        }

        if ($conflicts.Count -eq 0) {
            CreateNewAppointment -startTime $slot.Start -endTime $slot.End
            Write-Output "Created appointment from $($slot.Start.ToString()) to $($slot.End.ToString()) ($($interval) min)"
            return
        }
    }

    Write-Output "No available slot found on $($date.ToString("dddd, MMMM dd, yyyy"))"
}

Write-Output "Task: Hunting for free time.."

$today = Get-Date
$monday = $today.AddDays(-($today.DayOfWeek.value__ - 1))
$friday = $monday.AddDays(4)
for ($date = $monday; $date -le $friday; $date = $date.AddDays(1)) {
    $startOfDay = $date.Date.AddHours(10)
    $endOfDay = $date.Date.AddHours(16)

    FindFreeSlot -appointments $appointments -startDate $startOfDay -endDate $endOfDay
}