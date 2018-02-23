<#
.SYNOPSIS 
    HP Universal Print Driver Upgrade
.DESCRIPTION 
    Installs the latest HP UPD and upgrades any existing printers with the new driver.
.EXAMPLE 
    Upgrade-HPUPD
.NOTES 
    Name       : Upgrade-HPUPD
    Author     : Tom Dobson & Radoslav Radoev
    Version    : 1.1
    DateCreated: 02-12-2018
    DateUpdated: 02-15-2018
        Changes: Better existing driver install detection
                 Logging verbosity
                 Variable scoping changes

.LINK 
#>

#region Defining Functions...

# Write log file
Function Write-Log
{
       [CmdletBinding()]
       Param (
              [Parameter(Mandatory = $False, HelpMessage = "Log Level")]
              [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
              [string]$Level = "INFO",
              [Parameter(Mandatory = $True, Position = 0, HelpMessage = "Message to be written to the log")]
              [string]$Message,
              [Parameter(Mandatory = $False, HelpMessage = "Log file location and name")]
              [string]$Logfile = "C:\Temp\$DRIVERSTORE.log"
       )
    BEGIN {
       $Stamp = (Get-Date).toString("yyyy/MM/dd HH:mm:ss")
           $Line = "$Stamp $Level $Message`r`n"
    }
    PROCESS {
       If ($Logfile) {
            [System.IO.File]::AppendAllText($Logfile, $Line)
           } Else {
                  Write-Output $Line
           }
    }
    END {}
} # END Write-Log


Function Upgrade-HPUPD 
{

        # Variables and Constants
    if (-Not ($DRIVER)) 
    {
        New-Variable -Name "DRIVER" -Value "HP Universal Printing PCL 6 (v6.5.0),3,Windows x64" -Option Constant
    }

    if (-Not ($DRIVERSTORE)) 
    {
        New-Variable -Name "DRIVERSTORE" -Value "HP Universal Printing PCL 6 (v6.5.0)" -Option Constant
    }

  
    #Validates the script is being run as admin
    Function Check-AdminRights {
        $Wid=[System.Security.Principal.WindowsIdentity]::GetCurrent()
        $Prp=new-object System.Security.Principal.WindowsPrincipal($Wid)
        $Adm=[System.Security.Principal.WindowsBuiltInRole]::Administrator
        $IsAdmin=$prp.IsInRole($Adm)
        return $IsAdmin
    }

    Function Check-ForHPUPD {
             
        #Checks for Sharp UPD In Driver Store.
        $WIN7UPD = "$ENV:windir\System32\DriverStore\FileRepository\hpcu210u.inf_amd64_neutral_1f5cb1d03fe18675\hpcu210u.inf"
        $UPD = "$ENV:windir\System32\DriverStore\FileRepository\hpcu210u.inf_amd64_1f5cb1d03fe18675\hpcu210u.inf"
             
        If ((Test-Path $UPD) -or (Test-Path $WIN7UPD)) 
        {
            Write-Log "$DRIVERSTORE is already detected in the Driver Store."
            return $True
        } 
        Else 
        {
            Write-Log "Could not find $UPD." -Level WARN
            return $False
        }
    }

    Function Install-Driver {

        $infPath = "\\shcsd\sharp\drivers\Printers\HP\_HP Universal Print Driver 6.5.0.22695\Extracted\hpcu210u.inf"

        rundll32 printui.dll PrintUIEntry /ia /m $DRIVERSTORE /h "x64" /v "Type 3 - User Mode" /f $infPath
        Write-Log "Installing $DRIVERSTORE"
    
    }


    Function Add-UPDtoDriverStore {

        #Adds Sharp UPD to Driver Store if its not already there.
        $installedPrinterDrivers = gwmi Win32_PrinterDriver
      
        if (-Not (Check-ForHPUPD)) {
       
            Write-Log "Adding $DRIVERSTORE to Driver Store."
            Install-Driver

        } elseif ($installedPrinterDrivers.name -notcontains $DRIVER) {

            Install-Driver

        } else {

            Write-Log "$DRIVERSTORE is in DriverStore and is already Installed."

        }
        
        #manual driver removal : printui /s /t2
    }

    Function Upgrade-ToUPD {

        #Upgrades Existing Sharp Drivers
        $installedPrinters = Get-WmiObject Win32_Printer


        for ($i = 1; $i -le 10; $i++) {
            Write-Log "Waiting for driver to load into store. Waiting for: $($i * 5) seconds"
            if (gwmi win32_printerdriver | where {$_.Name -eq $DRIVER}) {break}
            Start-Sleep -Seconds 5
        }

        Write-Log "Checking HP Printers printers"
        foreach($printer in ($installedPrinters|Where{$_.DriverName -like 'HP Universal Printing PCL 6*'})){
            $name = $printer.name
            Write-Log "Upgrading $name to $DRIVERSTORE"
            & rundll32 printui.dll PrintUIEntry /Xs /n $name DriverName $DRIVERSTORE
        }
}

#endregion

#region Run Script   
    if (Check-AdminRights) {
        Add-UPDtoDriverStore
        Upgrade-ToUPD
    } else {
        $Message = "This script requires Admin Rights, please rerun as admin"
        Write-Log $Message -Level ERROR
        Write-Warning $Message
    }

#endregion   
}
