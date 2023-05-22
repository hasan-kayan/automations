$processName = Read-Host -Prompt "Please enter the process name you would like to check"

$process = Get-Process -Name $processName -ErrorAction SilentlyContinue

if ($process) {
    Write-Output "$processName program is running."

    if ($process.Responding) {
        Write-Output "$processName is responding."
    } else {
        Write-Output "$processName is not responding."
    }

    if ($process.Threads | Where-Object { $_.WaitReason }) {
        Write-Output "$processName is in a waiting state."
    } else {
        Write-Output "$processName is not in a waiting state."
    }

    if ($process.Threads | Where-Object { $_.ThreadState -eq "Suspended" }) {
        Write-Output "$processName is suspended."
    } else {
        Write-Output "$processName is not suspended."
    }
} else {
    Write-Output "$processName is not running."
}
