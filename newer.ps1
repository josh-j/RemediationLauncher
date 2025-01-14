# Script configuration
$script:UseParallelExecution = $true  # Toggle this to enable/disable parallel processing
$script:MaxConcurrentJobs = 10        # Maximum number of concurrent jobs when parallel execution is enabled

#Requires -RunAsAdministrator
#Requires -Version 5.0

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$RemediationName,
    
    [Parameter(Mandatory = $false)]
    [string]$RemediationsRoot = (Join-Path $PSScriptRoot "Remediations"),
    
    [Parameter(Mandatory = $false)]
    [string]$RemoteTempDir = 'C:\ProgramData\86CSRemediations',
    
    [Parameter(Mandatory = $false)]
    [int]$ConnectionTimeout = 30,
    
    [Parameter(Mandatory = $false)]
    [switch]$CleanupOnCompletion,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential = $null
)

Begin {
    # Initialize error handling
    $ErrorActionPreference = 'Stop'
    
    # Setup logging paths
    $script:MainLogPath = Join-Path $PSScriptRoot "Logs\$RemediationName-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
    $script:ErrorLog = Join-Path $PSScriptRoot "Logs\$RemediationName-Errors-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"
    $script:ResultsPath = Join-Path $PSScriptRoot "Results\$RemediationName-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv"

    # Ensure log directories exist
    @(Split-Path $script:MainLogPath), (Split-Path $script:ResultsPath) | ForEach-Object {
        if (-not (Test-Path $_)) {
            New-Item -ItemType Directory -Path $_ -Force | Out-Null
        }
    }

    function Write-Log {
        param(
            [string]$Message,
            [ValidateSet('Info', 'Warning', 'Error')]
            [string]$Level = 'Info'
        )
        
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logMessage = "[$timestamp] [$Level] $Message"
        
        if ($script:UseParallelExecution) {
            # Thread-safe logging for parallel execution
            $lockObject = [System.Threading.Mutex]::new($false, "Global\RemediationScriptLog")
            try {
                [void]$lockObject.WaitOne()
                switch ($Level) {
                    'Info' { Write-Host $logMessage -ForegroundColor Green }
                    'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
                    'Error' { Write-Host $logMessage -ForegroundColor Red }
                }
                Add-Content -Path $script:MainLogPath -Value $logMessage
                if ($Level -eq 'Error') {
                    Add-Content -Path $script:ErrorLog -Value $logMessage
                }
            }
            finally {
                if ($lockObject) {
                    $lockObject.ReleaseMutex()
                    $lockObject.Dispose()
                }
            }
        }
        else {
            # Standard logging for sequential execution
            switch ($Level) {
                'Info' { Write-Host $logMessage -ForegroundColor Green }
                'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
                'Error' { Write-Host $logMessage -ForegroundColor Red }
            }
            Add-Content -Path $script:MainLogPath -Value $logMessage
            if ($Level -eq 'Error') {
                Add-Content -Path $script:ErrorLog -Value $logMessage
            }
        }
    }

    function Invoke-Remediation {
        param(
            [Parameter(Mandatory = $true)]
            [string]$ComputerName
        )
        
        try {
            Write-Log "Processing computer: $ComputerName" -Level Info
            
            # Validate computer name format
            if ($ComputerName -match '[^\w\-\.]') {
                throw "Invalid computer name format: $ComputerName"
            }
            
            # Your existing remediation logic here
            # Make sure to properly escape any special characters
            # and validate all inputs before processing
            
            return @{
                ComputerName = $ComputerName
                Status = 'Success'
                Message = "Remediation completed successfully"
                Timestamp = Get-Date
            }
        }
        catch {
            Write-Log "Error processing $ComputerName`: $_" -Level Error
            return @{
                ComputerName = $ComputerName
                Status = 'Failed'
                Message = $_.Exception.Message
                Timestamp = Get-Date
            }
        }
    }
}

Process {
    try {
        Write-Log "Starting remediation deployment: $RemediationName" -Level Info
        Write-Log "Using remediation root: $RemediationsRoot" -Level Info
        Write-Log "Parallel execution: $(if ($script:UseParallelExecution) { 'Enabled' } else { 'Disabled' })" -Level Info
        
        $remediationPath = Join-Path $RemediationsRoot $RemediationName
        if (-not (Test-Path $remediationPath)) {
            throw "Remediation path not found: $remediationPath"
        }

        $computerListPath = if (Test-Path -Path (Join-Path $remediationPath 'failed_computers.txt')) {
            Write-Log "failed_computers.txt found... ignoring computers.txt" -Level Info
            Join-Path $remediationPath 'failed_computers.txt'
        } else {
            Join-Path $remediationPath 'computers.txt'
        }

        $computers = Get-Content $computerListPath -ErrorAction Stop | 
                    Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
                    ForEach-Object { $_.Trim() }
        
        Write-Log "Found $($computers.Count) computers to process" -Level Info

        if ($script:UseParallelExecution) {
            # Parallel processing implementation
            $results = $computers | ForEach-Object -ThrottleLimit $script:MaxConcurrentJobs -Parallel {
                $computer = $_
                try {
                    Invoke-Remediation -ComputerName $computer
                }
                catch {
                    @{
                        ComputerName = $computer
                        Status = 'Failed'
                        Message = $_.Exception.Message
                        Timestamp = Get-Date
                    }
                }
            }
        }
        else {
            # Sequential processing
            $results = $computers | ForEach-Object { 
                Invoke-Remediation -ComputerName $_
            }
        }

        # Export and process results
        $results | Export-Csv -Path $script:ResultsPath -NoTypeInformation -Force
        
        $successful = @($results | Where-Object Status -eq 'Success')
        $failed = @($results | Where-Object Status -eq 'Failed')
        
        Write-Log "Remediation complete. Success: $($successful.Count), Failed: $($failed.Count)" -Level Info
        
        if ($failed.Count -gt 0) {
            $failed.ComputerName | Set-Content -Path (Join-Path $remediationPath 'failed_computers.txt') -Force
            Write-Log "Failed computers written to: $(Join-Path $remediationPath 'failed_computers.txt')" -Level Warning
        }
    }
    catch {
        Write-Log "Critical error in main execution: $_" -Level Error
        throw
    }
}

End {
    if ($CleanupOnCompletion) {
        Write-Log "Cleaning up temporary files..." -Level Info
        # Add cleanup logic here
    }
    Write-Log "Script execution completed" -Level Info
}
