# Assign Script Title
$scriptName = $MyInvocation.MyCommand.Name
$Host.UI.RawUI.WindowTitle = "$scriptName - Created by David Letts"

# Global Variables
$hostname = hostname
$psexec = "C:\File\Path\Redacted\PsExec.exe"
$terminals = "$PSScriptRoot\terminal_list.txt"
$v314Installed = "$PSScriptRoot\v314Installed.txt"
$v336Installed = "$PSScriptRoot\v336Installed.txt"
$BothInstalled = "$PSScriptRoot\BothInstalled.txt"
$NoDrivers = "$PSScriptRoot\NoDrivers.txt"
$uname = "UsernameRedactedForSecurity"
$pword = "PasswordRedactedForSecurity"

# Delete existing files
Remove-Item $terminals -ErrorAction SilentlyContinue
Remove-Item $v314Installed -ErrorAction SilentlyContinue
Remove-Item $v336Installed -ErrorAction SilentlyContinue
Remove-Item $BothInstalled -ErrorAction SilentlyContinue
Remove-Item $NoDrivers -ErrorAction SilentlyContinue

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
Write-Host "                  Ingenico Driver Version Check (v3.14 or v3.36)                " -ForegroundColor Black -BackgroundColor Yellow
Write-Host "                             Created by: David Letts                            " -ForegroundColor Black -BackgroundColor Yellow
Write-Host ""
Write-Host "Notes:" -BackgroundColor Cyan -ForegroundColor Black
Write-Host "This script will scan all POS to check for the Ingenico USB Drivers, either v3.14 or v3.36." -ForegroundColor Cyan
Write-Host "This script is not compatible with kiosks as they are out of scope for this project." -ForegroundColor Cyan
Write-Host "Kiosk compatibility could be achieved for future projects." -ForegroundColor Cyan
Write-Host ""
Write-Host "Select which terminals to check from the window that will appear."
Pause
notepad $terminals
Write-Host ""
Read-Host "Press Enter to begin scanning POS at $hostname"

# Check driver versions
Write-Host "Checking driver versions..."

# Check if POS is online
Get-Content $terminals | ForEach-Object {
    $terminal = $_
    Write-Host "Testing connection to $terminal"
    $terminalOnline = Test-Connection -Computername $terminal -BufferSize 16 -Count 1 -Quiet

    if ($terminalOnline -eq $True) {
        Write-Host "$terminal is online." -ForegroundColor Green
        Write-Host ""

        $creds = New-Object System.Management.Automation.PSCredential("$terminal\$uname", $($pword | ConvertTo-SecureString -AsPlainText -Force))
        New-PSDrive -Name "z" -PSProvider FileSystem -Root "\\$terminal\C$" -Credential $creds | Out-Null

        # Check for v3.36
        if (Test-Path "z:\Path\Redacted\Uninstall.exe") {
            if (Test-Path "z:\Path\Redacted\Uninstall.exe") {
                Write-Host "Both v3.14 and v3.36 are currently installed on $terminal." -ForegroundColor Yellow
                Add-Content -Path $BothInstalled -Value $terminal
            } else {
                Write-Host "v3.36 is currently installed on $terminal."
                Add-Content -Path $v336Installed -Value $terminal
            }
        }

        # Check for v3.14
        elseif (Test-Path "z:\Path\Redacted\Uninstall.exe") {
            Write-Host "v3.14 is currently installed on $terminal."
            Add-Content -Path $v314Installed -Value $terminal

        # If no drivers were found
        } else {
            Write-Host "No drivers were found on $terminal."
            Add-Content -Path $NoDrivers -Value $terminal
        }

        Remove-PSDrive -Name "z"
        Get-PSSession -ComputerName $terminal -Credential $creds | Remove-PSSession
        Write-Host ""

    } else {
        Write-Host "$terminal is offline." -ForegroundColor Red
        Write-Host ""
    }
}

# Check if any POS are running old driver
if (Test-Path $v314Installed) {
    Write-Host ""
    Write-Host "There were terminals with the old driver (v3.14) installed." -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to see which POS need upgraded"
    notepad.exe $v314Installed

# Check if any POS are running both drivers
} if (Test-Path $BothInstalled) {
    Write-Host ""
    Write-Host "There were terminals with both drivers (v3.14 & v3.36) installed." -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to see which POS need upgraded"
    notepad.exe $BothInstalled

# Check if any POS are missing drivers
} if (Test-Path $NoDrivers) {
    Write-Host ""
    Write-Host "There were terminals with no drivers installed (v3.14 or v3.36)." -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to see which POS need upgraded"
    notepad.exe $NoDrivers

# If all POS are running new driver
} if (Test-Path $v336Installed) {
    if (!(Test-Path $v314Installed)) {
        if (!(Test-Path $BothInstalled)) {
            if (!(Test-Path $NoDrivers)) {
                Write-Host ""
                Write-Host "All POS are running new driver (v3.36)." -ForegroundColor Black -BackgroundColor Green
                Write-Host ""
            }
        }
    }

# If no drivers are installed
} else {
    Write-Host ""
    Write-Host "No drivers were found on any POS." -ForegroundColor Red
    Write-Host ""
}

# Display # of successes/failures
$results = @($v314Installed, $v336Installed, $BothInstalled, $NoDrivers)

foreach ($file in $results) {
    if (Test-Path $file) {
        $count = (Get-Content $file | Measure-Object -Line).Lines
        Write-Host "$file contains $count terminals" -ForegroundColor Cyan
        Write-Host ""
    }
}

# Pause at the end
Read-Host "Press Enter to exit"
