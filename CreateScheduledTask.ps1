<#
.SYNOPSIS
    Used for  firmware and drivers for RightSight connected to a MeetUp.
    Requires a file share that contains the Install-RightSight.ps1 and firmware file
    The account used in the Scheduled Task must have read access to the share file.
    
.DESCRIPTION
   This will create a scheduled task on a remote machine that will deploy the MeetUp firmware and the RightSight software.
   You must only pass the name of a computer, not the object, so you will need to use a select-object
   If the username or password entered for the account failes, the script will stop (since it will fail on remaining servers as well)

.NOTES
    File Name           : CreateScheduledTask.ps1
    Author              : Adam Berns (aberns@logitech.com)
    Prerequisite        : PowerShell V2 or later
    Script posted over  : Google Files (check with author for access)

.PARAMETER RightSight
    RightSight      : How Right Sight will be used, one of the following options
        Dynamic     : Will always run Dynamic Framing
        OCS         : On Call Start. It will frame only when the meeting starts. To reset to Dynaminc use the home button on the control
        Off         : No Framing

.PARAMETER RemoteComputer
    RemoteComputer  : Pulled from Pipeline

.PARAMETER SourceFolder
    SourceFolder    : Location that contains the firmware installer, and the install-rightsight.ps1 script

.PARAMETER username
    username        : The account used for the scheuduled task to run as. Thist must be a valid account

.PARAMETER StartTime
    If Omited the task can only be ran on-demand.
    StartTime       : This is a DateTime value of when you want the script to execute
    Examples        : Any of the formats from this output: (Get-Date).GetDateTimeFormats()

.EXAMPLE 
    "server1" | CreateScheduledTask.ps1 -SourceFolder "\\server\share\fwupdate\" -UserName "domain\usern" -RightSight Dynamic -StartTime "February 12, 2019 12:15:00 PM"
    get-adcomputer -SearchBase "CN=Computers,DC=domain,DC=com" | select-object name |  CreateScheduledTask.ps1 -SourceFolder "\\server\share\fwupdate\" -UserName "domain\usern" -RightSight Dynamic -StartTime "February 12, 2019 12:15:00 PM"
    import-csv serverlist.csv |  select-object remotecomputername | CreateScheduledTask.ps1 -SourceFolder "\\server\share\fwupdate\" -UserName "domain\usern" -RightSight Dynamic
#>

[CmdletBinding(SupportsShouldProcess)]
Param (
    [Parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName)]
    [array]$remotecomputer,

    [Parameter(Mandatory=$true,ValueFromPipeline=$false)]
    [string]$SourceFolder,

    [Parameter(Mandatory=$true,ValueFromPipeline=$false)]
    [string]$username,

    [Parameter(Mandatory=$true,ValueFromPipeline=$false)]
    [ValidateSet("off","Dynamic","OCS")]
    [Alias("RS")]
    [string]$RightSight,

    [Parameter(Mandatory=$false,ValueFromPipeline=$false)]
    [datetime]$StartTime

)


Begin {

    $Installarg = '-NonInteractive -NoLogo -NoProfile -ExecutionPolicy ByPass -file "%temp%\fwupdate\Install-rightsight.ps1" -rightsight ' + $RightSight
    $copyargs = $Sourcefolder  + " %temp%\fwupdate\"

    $taskname = "Install RightSight and MeetUp firmware"
    $Description = "Installs RightSight and Meetup Firmware"
    if (-not $password) {
        $secpassword = read-host -AsSecureString -Prompt "Enter Password for $username"
        $secpassword2 = read-host -AsSecureString -Prompt "Confirm Password for $username"

        $pwd1_text = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secpassword))
        $pwd2_text = [Runtime.InteropServices.Marshal]::PtrToStringAuto([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secpassword2))


        if ($pwd1_text -ceq $pwd2_text) {
            $Credentials = New-Object System.Management.Automation.PSCredential -ArgumentList $UserName, $SecPassword
            $Password = $Credentials.GetNetworkCredential().Password
        } else {
            write-host "Passwords do not match"
            exit
        }
    }
}

process {
    $ErrorActionPreference = "Stop"
    write-host "processing computer: $remotecomputer"
    $error.Clear()
    try {
        $Action = (New-ScheduledTaskAction -Execute "xcopy" -Argument $copyargs -CimSession $remotecomputer), (New-ScheduledTaskAction -execute 'PowerShell.exe' -Argument $Installarg  -CimSession $remotecomputer)
        $setting = New-ScheduledTaskSettingsSet -ExecutionTimeLimit (new-timespan -Minutes 5)  -CimSession $remotecomputer

        if ($starttime) {  
            $trigger = New-ScheduledTaskTrigger -Once -At $starttime  -CimSession $remotecomputer
            $task = New-ScheduledTask -Action $action   -Settings $setting -Description $Description -Trigger $trigger  -CimSession $remotecomputer

        }
        else {
            $task = New-ScheduledTask -Action $action   -Settings $setting -Description $Description  -CimSession $remotecomputer
        }
            
        Register-ScheduledTask -InputObject $task -TaskName $taskname -User $username -Password $Password  -CimSession $remotecomputer -ErrorAction stop
    }

    catch {
        Switch ([string]$error[0].Exception.MessageID) {
        "HRESULT 0x8007052e" {
            write-host "Password enterned for the account is not correct, Exiting" -ForegroundColor Red
            exit
            }
        "HRESULT 0x80070534" {
            write-host "Username enterned for the account is not correct, Exiting" -ForegroundColor Red
            exit
            }
        "HRESULT 0x800700b7" {
            write-host "Scheduled Task Already Exists (continuing)" -ForegroundColor yellow
            }

    Default {
        write-host "Continuing Error : $Error[0].Exception (continuing)" -ForegroundColor Yellow
    }

    }
    
    }#End Catch
}

end {
    write-host "Finished proccesing"
}
