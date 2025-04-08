try {
    $username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $cleanUsername = $username.Split('\')[1] 

    if ($cleanUsername -match "\.") {
        $USERPC = "$cleanUsername@example.com"
    } elseif ($cleanUsername -match "-") {
        $USERPC = $cleanUsername -replace "-", "."
        $USERPC = "$USERPC@example.com"
    } else {
        Write-Host "No matching email found. Please provide a new PC name using '-' instead of '.' (e.g., john-doe):"
        $NEWUSERNAME = Read-Host "Enter new PC name"
        Write-Host "Renaming computer to: $NEWUSERNAME ..."
        exit
    }

    Write-Host "Detected Email: $USERPC"

    $csvFile = "C:\Path\To\Your\CSV\File.csv"
    if (-Not (Test-Path $csvFile)) {
        throw "CSV file not found at: $csvFile"
    }

    $csvData = Import-Csv -Delimiter ";" -Path $csvFile
    $emailHeader = if ($csvData[0].PSObject.Properties["email"]) { "email" } elseif ($csvData[0].PSObject.Properties["emai"]) { "emai" } else { throw "CSV does not contain an 'email' column." }

    $userRow = $csvData | Where-Object { $_.$emailHeader.Trim().ToLower() -eq $USERPC.ToLower() }

    if (-not $userRow) {
        throw "No matching email found in CSV. Please restart script and provide a valid PC name."
    }

    $email = $userRow.$emailHeader.Trim()
    $password = $userRow.password.Trim()

    Write-Host "Matching credentials found: $email | (Password is hidden for security)"

    $outlookPath = "C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE"
    if (-Not (Test-Path $outlookPath)) {
        throw "Outlook executable not found at: $outlookPath"
    }

    Write-Host "Launching Outlook..."
    Start-Process -FilePath $outlookPath
    Start-Sleep -Seconds 3

    $wshell = New-Object -ComObject WScript.Shell
    Start-Sleep -Seconds 3
    $wshell.SendKeys("{ENTER}")

    Start-Sleep -Seconds 5
    $wshell.SendKeys($email)
    Start-Sleep -Seconds 1
    $wshell.SendKeys("{TAB}")
    Start-Sleep -Seconds 1
    $wshell.SendKeys("{TAB}")
    Start-Sleep -Seconds 1
    $wshell.SendKeys("{ENTER}")
    Start-Sleep -Seconds 8

    $portsDetected = @()
    $netstatOutput = netstat -ano | Select-String -Pattern "OUTLOOK"
    $netstatOutput | ForEach-Object {
        if ($_ -match "\s+\S+\s+(\S+):(\d+)\s+") {
            $localPort = $matches[2]
            $portsDetected += $localPort
        }
    }
    if ($portsDetected.Count -eq 5) {
        Write-Host "5 ports detected, waiting 5 seconds before executing password section..."
        Start-Sleep -Seconds 5
    }

    $wshell.SendKeys($password)
    Start-Sleep -Seconds 1
    $wshell.SendKeys("{ENTER}")
    Start-Sleep -Seconds 5
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys("{TAB}")
    $wshell.SendKeys("{ENTER}")

    Write-Host "Credentials submitted."
}
catch {
    Write-Error "An error occurred: $_"
}
