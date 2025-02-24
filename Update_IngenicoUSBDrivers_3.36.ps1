# Assign Script Title
$scriptName = $MyInvocation.MyCommand.Name
$Host.UI.RawUI.WindowTitle = "$scriptName - Created by David Letts"

# Global Variables
$hostname = hostname
$psexec = "C:\File\Path\Redacted\PsExec.exe"
$terminals = "$PSScriptRoot\terminal_list.txt"
$terminalsOnline = "$PSScriptRoot\terminals_online.txt"
$removalPassed = "$PSScriptRoot\removal_passed.txt"
$removalFailed = "$PSScriptRoot\removal_failed.txt"
$installPassed = "$PSScriptRoot\install_passed.txt"
$installFailed = "$PSScriptRoot\install_failed.txt"
$uname = "UsernameRedactedForSecurity"
$pword = "PasswordRedactedForSecurity"

# Delete existing files
Remove-Item $removalPassed -ErrorAction SilentlyContinue
Remove-Item $removalFailed -ErrorAction SilentlyContinue
Remove-Item $installPassed -ErrorAction SilentlyContinue
Remove-Item $installFailed -ErrorAction SilentlyContinue
Remove-Item $terminals -ErrorAction SilentlyContinue
Remove-Item $terminalsOnline -ErrorAction SilentlyContinue

Clear-Host

# Generate terminal_list.txt
function Get-TerminalByType {
    param (
        $Type
    )
    <#
    Available types
    -------------------
    ALL
    BAR
    BOX
    CONC
    DUAL
    KIOSK
    #>

    $Query = "
        SET NOCOUNT ON
        SELECT terminal_name
        FROM table_name_redacted
        WHERE timestamp = (
            SELECT MAX(timestamp) 
            FROM table_name_redacted
        )
        AND terminal_name NOT LIKE '%admin%'
    "

    $AllTerminals = sqlcmd -E -d "different_table_name_redacted" -Q $Query -h-1
    $AllTerminals = $AllTerminals.Trim()

    switch ($Type) {
        "ALL"   {$FilteredTerminals = $AllTerminals                                    }
        "BAR"   {$FilteredTerminals = $AllTerminals | Where-Object{$_ -like "*BAR*"}   }
        "BOX"   {$FilteredTerminals = $AllTerminals | Where-Object{$_ -like "*BOX*"}   }
        "CONC"  {$FilteredTerminals = $AllTerminals | Where-Object{$_ -like "*CONC*"}  }
        "DUAL"  {$FilteredTerminals = $AllTerminals | Where-Object{$_ -like "*DUAL*"}  }
        "KIOSK" {$FilteredTerminals = $AllTerminals | Where-Object{$_ -like "*KIOSK*"} }
        Default {$FilteredTerminals = $AllTerminals                                    }
    }

    return $FilteredTerminals
}

$terminal_List = @()
$terminal_List += Get-TerminalByType "BAR"
$terminal_List += Get-TerminalByType "BOX"
$terminal_List += Get-TerminalByType "CONC"
$terminal_List += Get-TerminalByType "DUAL"
$terminal_List | Out-File $terminals -Encoding ascii

#Script Header
Write-Host ""
Write-Host "                                                                                " -BackgroundColor Red
Write-Host "                                                                                "
Write-Host "            ████████████      ██████████  ██████████          ████████████      " -ForegroundColor Red
Write-Host "        ████████    ████    ████      ██████      ████      ████        ████    " -ForegroundColor Red
Write-Host "      ██████        ████  ████          ████        ████    ██            ███   " -ForegroundColor Red
Write-Host "    ████            ████  ████          ██          ████    ██                  " -ForegroundColor Red
Write-Host "    ██              ████  ████          ██          ████    ██                  " -ForegroundColor Red
Write-Host "  ████              ████  ████          ██          ████    ██            ███   " -ForegroundColor Red
Write-Host "  ████              ████  ████          ██          ████    ████        ████    " -ForegroundColor Red
Write-Host "  ████████████████  ████  ████          ██          ████      ████████████      " -ForegroundColor Red
Write-Host "                                                                                "
Write-Host "                                                                                " -BackgroundColor Red
Write-Host ""
Write-Host "                      Ingenico Driver v3.36 Upgrade Project                     " -ForegroundColor Black -BackgroundColor Yellow
Write-Host "                             Created by: David Letts                            " -ForegroundColor Black -BackgroundColor Yellow
Write-Host ""
Write-Host "Notes:" -BackgroundColor Cyan -ForegroundColor Black
Write-Host "This script will automatically update the Ingenico USB Drivers from v3.14 to v3.36." -ForegroundColor Cyan
Write-Host "This script is not compatible with kiosks as they are out of scope for this project." -ForegroundColor Cyan
Write-Host "Kiosk compatibility could be achieved for future projects." -ForegroundColor Cyan
Write-Host ""
Write-Host "Select which terminals to update from the window that will appear."
Pause
notepad $terminals
Write-Host ""
Read-Host "Press Enter to begin upgrade at $hostname"


# Copy "IngenicoUSBDrivers_3.36" folder to C:\Path\Redacted on MAIN if missing
if (-not (Test-Path "C:\Path\Redacted\IngenicoUSBDrivers_3.36")) {
    Write-Host ""
    Write-Host "Could not find 'C:\Path\Redacted\IngenicoUSBDrivers_3.36' on $hostname" -ForegroundColor Red
    Write-Host "Copying '$PSScriptRoot\Out\IngenicoUSBDrivers_3.36' to 'C:\Path\Redacted\IngenicoUSBDrivers_3.36'"
    Copy-Item "$PSScriptRoot\Out\IngenicoUSBDrivers_3.36" "C:\Path\Redacted\IngenicoUSBDrivers_3.36" -Recurse -Force
    Write-Host "Copied 'IngenicoUSBDrivers_3.36' to $hostname successfully" -ForegroundColor Green
}

Write-Host ""
Write-Host "Copying files to selected terminals..."
Write-Host ""

# Copy "IngenicoUSBDrivers_3.36" folder to all POS in terminal_list.txt
$folder = "C:\Path\Redacted\IngenicoUSBDrivers_3.36"
Get-Content $terminals | ForEach-Object {
    $terminal = $_
    Write-Host "Testing connection to $terminal"
    $terminalOnline = Test-Connection -Computername $terminal -BufferSize 16 -Count 1 -Quiet

    if ($terminalOnline -eq $True) {
        Write-Host "$terminal is online." -ForegroundColor Green
        Add-Content -Path $terminalsOnline -Value $terminal

        $creds = New-Object System.Management.Automation.PSCredential("$terminal\$uname", $($pword | ConvertTo-SecureString -AsPlainText -Force))
        New-PSDrive -Name "z" -PSProvider FileSystem -Root "\\$terminal\C$" -Credential $creds | Out-Null

        Write-Host "Copying $folder to $terminal"
        Copy-Item "$folder" -Destination "\\$terminal\c$\Path\Redacted\IngenicoUSBDrivers_3.36" -Recurse

        Remove-PSDrive -Name "z"
        Get-PSSession -ComputerName $terminal -Credential $creds | Remove-PSSession
    } else {
        Write-Host "$terminal is offline." -ForegroundColor Red
    }
}

# Uninstall old drivers & reboot POS
Get-Content $terminalsOnline | ForEach-Object {
    $terminal = $_
    New-PSDrive -Name "z" -PSProvider FileSystem -Root "\\$terminal\C$" -Credential $creds | Out-Null

    # Check if Uninstall.exe exists, if not add to removal_failed.txt
    if (-not (Test-Path "z:\Path\Redacted\IngenicoUSBDrivers\Uninstall.exe")) {
        Write-Host ""
        Write-Host "Could not find 'C:\Path\Redacted\IngenicoUSBDrivers\Uninstall.exe' on $terminal." -ForegroundColor Red
        Write-Host "Unable to remove IngenicoUSBDrivers_3.14, adding to manual installation list." -ForegroundColor Red
        Add-Content -Path $removalFailed -Value $terminal
    } else {
        Add-Content -Path $removalPassed -Value $terminal

        # Uninstall old drivers
        $uninstallCommand = "$psexec \\$terminal /accepteula /u $uname /p $pword cmd /c `"C:\Path\Redacted\IngenicoUSBDrivers\Uninstall.exe`" /S"
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c", $uninstallCommand -NoNewWindow -Wait

        Write-Host ""
        Write-Host "Uninstalled IngenicoUSBDrivers_3.14 on $terminal, rebooting POS."
        Write-Host ""
        Start-Process -FilePath $psexec -ArgumentList "\\$terminal", "/accepteula", "/u", $uname, "/p", $pword, "cmd", "/c", "shutdown /r /t 0" -NoNewWindow -Wait

        Remove-PSDrive -Name "z"
    }
}

# Wait 20 seconds for terminal to shutdown to avoid conflict with next steps (e.g. installing new drivers before reboot)
Write-Host ""
Write-Host "Waiting 20 seconds for terminals to reboot..." -ForegroundColor Yellow
Timeout /t 20 /nobreak
Write-Host ""

# Install new drivers & reboot POS
foreach ($terminal in Get-Content $removalPassed) {
    # Check if POS is back online after 1st reboot
    $wait_time = 0
    $time_to_wait = 180
    while ($wait_time -lt $time_to_wait) {
        if (Test-Connection -ComputerName $terminal -Count 1 -Quiet) {
            Write-Host "$terminal is online." -ForegroundColor Green
            Write-Host ""
            Write-Host "Installing IngenicoUSBDrivers_3.36 on $terminal..."
            & $psexec \\$terminal /accepteula /u $uname /p $pword cmd /c "C:\Path\Redacted\IngenicoUSBDrivers_3.36\IngenicoUSBSilentInstall_336.cmd"
            Write-Host ""
            Write-Host "Rebooting $terminal..."
            & $psexec \\$terminal /accepteula /u $uname /p $pword cmd /c shutdown /r /t 0
            Add-Content -Path $installPassed -Value $terminal
            break
        }
        else {
            Write-Host ""
            Write-Host ""
            Write-Host "$terminal is still rebooting. Checking again in 5 seconds." -ForegroundColor Yellow
            Timeout /t 5 /nobreak
            $wait_time += 5
        }
    }

    # If POS does not come online in 3 minutes add to install_failed.txt and skip terminal
    if ($wait_time -ge $time_to_wait) {
        Write-Host "$terminal failed to come online after 3 minutes, adding to manual installation list" -ForegroundColor Red
        Add-Content -Path $installFailed -Value $terminal
    }
}

# Check for failed removals
if (Test-Path $removalFailed) {
    Write-Host ""
    Write-Host "There were terminals that failed to remove 'IngenicoUSBDriver v3.14'." -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to see which POS need manual installation"
    notepad.exe $removalFailed
}

# Check for failed reinstallations
if (Test-Path $installFailed) {
    Write-Host ""
    Write-Host "There were terminals that failed to upgrade the Ingenico USB Driver to v3.36." -ForegroundColor Red
    Write-Host "These terminals removed 'IngenicoUSBDriver v3.14' successfully, and only need v3.36 installed & a reboot." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to see which POS need manual installation"
    notepad.exe $installFailed
}

# All selected POS have successfully upgraded
if ($installPassed) {
    Write-Host ""
    Write-Host "The POS have finished upgrading the Ingenico USB Driver to v3.36." -ForegroundColor Black -BackgroundColor Green
    Write-Host "Please allow the POS to finish rebooting before testing." -ForegroundColor Yellow
    Write-Host "Verify functionality with site." -ForegroundColor Yellow
    Write-Host ""
}

# Display # of successes/failures
$results = @($removalFailed, $installFailed, $installPassed)

foreach ($file in $results) {
    if (Test-Path $file) {
        $count = (Get-Content $file | Measure-Object -Line).Lines
        Write-Host "$file contains $count terminals" -ForegroundColor Cyan
        Write-Host ""
    }
}

# Pause at the end
Read-Host "Press Enter to exit"
