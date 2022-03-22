# Log types to backup
$logArray = @("System","Security","Application")

# Name of month for folder name 
$currentMonth = Get-Date -UFormat %m
$month = (Get-Culture).DateTimeFormat.GetAbbreviatedMonthName($currentMonth)

# Path to Output
$destinationpath = "C:\Resultant\Logs\$month\"

# If the destination path does not exist it will create it
if (!(Test-Path -Path $destinationpath))
{
    New-Item -ItemType directory -Path $destinationpath
}

# Get the current date
$logdate = Get-Date -format yyyy-MM-dd-HHmm

# Start Process Timer
$StopWatch = [system.diagnostics.stopwatch]::startNew()

# Start Code
Clear-Host


Foreach($log in $logArray)
{
    # If using Clear and backup
    $destination = $destinationpath + $log + "-" + $logdate + ".evtx"

    Write-Host "Extracting the $log file now."

    # Extract each log file listed in $logArray from the local server.
    wevtutil epl $log $destination

    Write-Host "Clearing the $log file now."

    # Clear the log and backup to file.
    WevtUtil cl $log

}


# End Code

# Stop Timer
$StopWatch.Stop()
$TotalTime = $StopWatch.Elapsed.TotalSeconds
$TotalTime = [math]::Round($totalTime, 2)
write-host "The Script took $TotalTime seconds to execute."