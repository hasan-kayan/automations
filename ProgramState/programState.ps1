$process = Get-Process -Name "processName" -ErrorAction SilentlyContinue
if ($process) {
    Write-Output "$($process.ProcessName) programı çalışıyor."
    if ($process.Responding) {
        Write-Output "$($process.ProcessName) yanıt veriyor."
    } else {
        Write-Output "$($process.ProcessName) yanıt vermiyor."
    }
    if ($process.Threads | Where-Object { $_.WaitReason }) {
        Write-Output "$($process.ProcessName) bekleme durumunda."
    } else {
        Write-Output "$($process.ProcessName) bekleme durumunda değil."
    }
    if ($process.Threads | Where-Object { $_.ThreadState -eq "Suspended" }) {
        Write-Output "$($process.ProcessName) askıya alınmış durumda."
    } else {
        Write-Output "$($process.ProcessName) askıya alınmamış durumda."
    }
} else {
    Write-Output "$processName çalışmıyor."
}
