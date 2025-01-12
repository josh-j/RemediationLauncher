#Requires -Version 5.0

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$RemediationName,

    [Parameter(Mandatory = $true)]
    [string]$LogPath,

    [Parameter(Mandatory = $true)]
    [string]$RemediationDir
)

Begin {
    $ErrorActionPreference = 'Stop'
    # New-Item -ItemType Directory -Path (Split-Path $logPath) -Force -ErrorAction SilentlyContinue | Out-Null
    Start-Transcript -Path $LogPath -Force
}
Process {
    try {

        Add-Type -AssemblyName PresentationFramework
        Add-Type -AssemblyName System.Windows.Forms
        $Result = [System.Windows.Forms.MessageBox]::Show("Press ok for another test", "Test", 1)

        if ($Result -eq "OK") {
            msg * "test" "test"
        }
        else {
            msg * "test2" "test2"
        }
    }
    catch {
        Write-Error "Error executing $TaskName : $_"
        exit 1
    }
    finally {
        Stop-Transcript
    }
}
