#Requires -RunAsAdministrator
#Requires -Version 5.0

[CmdletBinding()]
param (
    [Parameter(Mandatory = $true)]
    [string]$RemediationName,

    [Parameter(Mandatory = $false)]
    [string]$RemediationsRoot = 'C:\Users\administrator\Projects\RemediationLauncher\Remediations',

    [Parameter(Mandatory = $false)]
    [string]$RemoteTempDir = 'C:\ProgramData\86CSRemediations',

    [Parameter(Mandatory = $false)]
    [int]$ConnectionTimeout = 30,

    [Parameter(Mandatory = $false)]
    [switch]$CleanupOnCompletion,

    [Parameter(Mandatory = $false)]
    [System.Management.Automation.PSCredential]$Credential = $null,

    [Parameter(Mandatory = $false)]
    [int]$MaxParallelJobs = 10  # New parameter for controlling parallel execution
)

Begin {
    Set-StrictMode -Version Latest
    $ErrorActionPreference = 'Stop'
    
    # Script-level paths
    $script:RemediationPath = Join-Path $RemediationsRoot $RemediationName
    $script:LogFolder = Join-Path $script:RemediationPath 'Logs'
    $script:DateStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $script:MainLogPath = Join-Path $script:LogFolder "Deployment_$($script:DateStamp).log"
    $script:ResultsPath = Join-Path $script:LogFolder "Results_$($script:DateStamp).csv"
    $script:ErrorLog = Join-Path $script:LogFolder "Errors_$($script:DateStamp).log"

    # Create log directory if it doesn't exist
    if (!(Test-Path $script:LogFolder)) {
        New-Item -ItemType Directory -Path $script:LogFolder -Force | Out-Null
    }

    try { 
        Start-Transcript -Path $script:MainLogPath -Force
        $script:TranscriptStarted = $true
    } 
    catch { 
        Write-Warning "Could not start transcript: $($_.Exception.Message)" 
        $script:TranscriptStarted = $false
    }

    function Write-Log {
        param([string]$Message, [ValidateSet('Info', 'Warning', 'Error')][string]$Level = 'Info')
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logMessage = "[$timestamp] [$Level] $Message"
        switch ($Level) {
            'Info' { Write-Host $logMessage -ForegroundColor Green }
            'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
            'Error' { Write-Host $logMessage -ForegroundColor Red }
        }
    }

    function Test-RemediationStructure {
        if (!(Test-Path $script:RemediationPath)) {
            throw "Remediation folder not found: $($script:RemediationPath)"
        }

        $computerListPath = Join-Path $script:RemediationPath 'computers.txt'
        $failedComputerListPath = Join-Path $script:RemediationPath 'failed_computers.txt'
        if (!(Test-Path $computerListPath) -and !(Test-Path $failedComputerListPath)) {
            throw 'computers.txt not found in remediation folder'
        }

        $foundScripts = @('AdminExec.ps1', 'CurrentUserExec.ps1', 'EveryUserExec.ps1', 'SystemExec.ps1') |
            Where-Object { Test-Path (Join-Path $script:RemediationPath $_) } |
            ForEach-Object { Write-Log "Found execution script: $_" -Level Info; $true } |
            Test-Any

        if (!$foundScripts) {
            throw "No execution scripts found"
        }
    }

    function Test-Any { 
        process { return $true } 
        end { return $false } 
    }

    function Invoke-ParallelRemediation {
        param(
            [Parameter(Mandatory)]
            [string[]]$ComputerNames
        )

        # Initialize throttle limit for parallel processing
        $throttleLimit = [math]::Min($MaxParallelJobs, $ComputerNames.Count)
        Write-Log "Starting parallel remediation with throttle limit: $throttleLimit" -Level Info

        # Create a thread-safe queue for results
        $resultQueue = [System.Collections.Concurrent.ConcurrentQueue[object]]::new()
        
        # Create runspace pool
        $runspacePool = [runspacefactory]::CreateRunspacePool(1, $throttleLimit)
        $runspacePool.Open()

        $runspaces = @()

        foreach ($computerName in $ComputerNames) {
            $powerShell = [powershell]::Create().AddScript({
                param($ComputerName, $RemediationPath, $RemoteTempDir, $RemediationName)
                
                try {
                    # Create session with increased timeout
                    $sessionOption = New-PSSessionOption -OpenTimeout 30000
                    $session = New-PSSession -ComputerName $ComputerName -SessionOption $sessionOption
                    
                    if (-not $session) {
                        throw "Failed to create PS Session"
                    }

                    $remotePath = Join-Path $RemoteTempDir $RemediationName

                    # Setup remote directory
                    Invoke-Command -Session $session -ScriptBlock { 
                        param($Path) 
                        if (-not (Test-Path $Path)) { 
                            New-Item -Path $Path -ItemType Directory -Force 
                        } 
                    } -ArgumentList $remotePath

                    # Copy files
                    Get-ChildItem $RemediationPath | 
                        Where-Object { $_.Name -notin @('computers.txt', 'Logs') } | 
                        Copy-Item -Destination $remotePath -ToSession $session -Recurse -Force

                    # Execute remediation (previous execution logic)
                    $results = Invoke-Command -Session $session -ScriptBlock {
                        param($remotePath, $RemediationName)
                
                        $results = @()
                        
                        # Execute AdminExec.ps1
                        $adminScriptPath = Join-Path $remotePath "AdminExec.ps1"
                        if (Test-Path $adminScriptPath) {
                            try {
                                & $adminScriptPath -RemediationName $RemediationName `
                                    -LogPath "$remotePath\admin_result.log" `
                                    -RemediationDir $remotePath
                                $results += @{
                                    Context = 'AdminExec'
                                    Status = 'Success'
                                }
                            }
                            catch {
                                $results += @{
                                    Context = 'AdminExec'
                                    Status = 'Failed'
                                    Error = $_.ToString()
                                }
                            }
                        }
        
                        # Execute SystemExec.ps1
                        $systemScriptPath = Join-Path $remotePath "SystemExec.ps1"
                        if (Test-Path $systemScriptPath) {
                            try {
                                Write-Host "Setting up SystemExec task"
                                $taskName = "Remediation_${RemediationName}_System"
                                
                                $action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
                                    -Argument "-ExecutionPolicy Bypass -NoProfile -File `"$systemScriptPath`" -RemediationName `"$RemediationName`" -LogPath `"$remotePath\system_result.log`" -RemediationDir `"$remotePath`""
                                
                                $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date)
                                $principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
                                $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
                                
                                Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force
                                Start-ScheduledTask -TaskName $taskName
                                
                                # Wait for completion
                                $timeout = (Get-Date).AddMinutes(5)
                                do {
                                    Start-Sleep -Seconds 2
                                    $taskInfo = Get-ScheduledTask -TaskName $taskName
                                    if ($taskInfo.State -eq 'Ready') { break }
                                } while ((Get-Date) -lt $timeout)
                                
                                $taskResult = Get-ScheduledTaskInfo -TaskName $taskName
                                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                                
                                if ($taskResult.LastTaskResult -eq 0) {
                                    $results += @{
                                        Context = 'SystemExec'
                                        Status = 'Success'
                                    }
                                } else {
                                    throw "Task failed with result: $($taskResult.LastTaskResult)"
                                }
                            }
                            catch {
                                $results += @{
                                    Context = 'SystemExec'
                                    Status = 'Failed'
                                    Error = $_.ToString()
                                }
                            }
                        }
        
                        # Execute CurrentUserExec.ps1
                        $currentUserScriptPath = Join-Path $remotePath "CurrentUserExec.ps1"
                        if (Test-Path $currentUserScriptPath) {
                            try {
                                Write-Host "Starting CurrentUserExec process"
                                
                                # Get active user using quser
                                Write-Host "Attempting quser fallback..."
                                $quserOutput = quser 2>&1 | Out-String
                                Write-Host "Quser output: $quserOutput"
                                
                                # Parse quser output
                                $currentUser = $null
                                $lines = $quserOutput -split '\r?\n' | Where-Object { $_ -match '\S' }
                                if ($lines.Count -gt 1) {  # First line is header
                                    $userLines = $lines | Select-Object -Skip 1
                                    foreach ($line in $userLines) {
                                        if ($line -match 'Active') {
                                            # Extract username which is always the first field
                                            $fields = $line.Trim() -split '\s+'
                                            $currentUser = $fields[0].Trim()
                                            Write-Host "Found active user from quser: $currentUser"
                                            break
                                        }
                                    }
                                }
                                
                                if (-not $currentUser) {
                                    throw "No active user detected"
                                }
                                
                                # Validate user
                                $validUserPattern = '^(?!(Public|defaultuser0|Default|.*\.adw|.*\.adf|Administrator|USAF_Admin)$)'
                                Write-Host "Checking if user '$currentUser' matches pattern: $validUserPattern"
                                if (-not ($currentUser -match $validUserPattern)) {
                                    throw "User $currentUser is not valid for execution"
                                }
                                
                                Write-Host "Checking user profile path: C:\Users\$currentUser"
                                if (-not (Test-Path "C:\Users\$currentUser")) {
                                    throw "User profile not found: $currentUser"
                                }
        
                                $taskName = "Remediation_${RemediationName}_$currentUser"
                                Write-Host "Creating scheduled task: $taskName"
                                
                                $action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
                                    -Argument "-ExecutionPolicy Bypass -WindowStyle Normal -NoProfile -File `"$currentUserScriptPath`" -RemediationName `"$RemediationName`" -LogPath `"$remotePath\current_user_result.log`" -RemediationDir `"$remotePath`""
                                
                                $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date)
                                $principal = New-ScheduledTaskPrincipal -UserId $currentUser -LogonType Interactive -RunLevel Highest
                                $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
                                
                                Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force
                                Start-ScheduledTask -TaskName $taskName
                                
                                $timeout = (Get-Date).AddMinutes(5)
                                do {
                                    Start-Sleep -Seconds 2
                                    $taskInfo = Get-ScheduledTask -TaskName $taskName
                                    Write-Host "Current task state: $($taskInfo.State)"
                                    if ($taskInfo.State -eq 'Ready') { break }
                                } while ((Get-Date) -lt $timeout)
                                
                                $taskResult = Get-ScheduledTaskInfo -TaskName $taskName
                                Write-Host "Task last result: $($taskResult.LastTaskResult)"
                                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                                
                                Write-Host "Task completed with result: $($taskResult.LastTaskResult)"
                                
                                # Create result object and convert to JSON
                                $resultObj = [PSCustomObject]@{
                                    Context = 'CurrentUserExec'
                                    Status = $(if ($taskResult.LastTaskResult -eq 0) { 'Success' } else { 'Failed' })
                                    Error = $(if ($taskResult.LastTaskResult -ne 0) { "Task failed with result: $($taskResult.LastTaskResult)" } else { $null })
                                }
                                
                                $jsonResult = $resultObj | ConvertTo-Json -Compress
                                Write-Host "Returning JSON result: $jsonResult"
                                
                                # Return just the JSON string
                                Write-Output $jsonResult
                            }
                            catch {
                                Write-Host "Error in CurrentUserExec: $_"
                                Write-Host "Error details: $($_.Exception)"
                                Write-Host "Stack trace: $($_.ScriptStackTrace)"
                                $results += @{
                                    Context = 'CurrentUserExec'
                                    Status = 'Failed'
                                    Error = $_.ToString()
                                }
                            }
                        }
        
                        # Execute EveryUserExec.ps1
                        $everyUserScriptPath = Join-Path $remotePath "EveryUserExec.ps1"
                        if (Test-Path $everyUserScriptPath) {
                            try {
                                Write-Host "Starting EveryUserExec process"
                                $validUserPattern = '^(?!(Public|defaultuser0|Default|.*\.adw|.*\.adf|Administrator|USAF_Admin)$)'
                                
                                $userProfiles = Get-ChildItem 'C:\Users' | 
                                    Where-Object { $_.PSIsContainer -and $_.Name -match $validUserPattern }
                                
                                foreach ($userProfile in $userProfiles) {
                                    $userName = $userProfile.Name
                                    Write-Host "Processing user: $userName"
                                    
                                    try {
                                        $taskName = "Remediation_${RemediationName}_$userName"
                                        
                                        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
                                            -Argument "-ExecutionPolicy Bypass -WindowStyle Normal -NoProfile -File `"$everyUserScriptPath`" -RemediationName `"$RemediationName`" -LogPath `"$remotePath\${userName}_result.log`" -RemediationDir `"$remotePath`""
                                        
                                        $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date)
                                        $principal = New-ScheduledTaskPrincipal -UserId $userName -LogonType Interactive -RunLevel Highest
                                        $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 5)
                                        
                                        Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force
                                        Start-ScheduledTask -TaskName $taskName
                                        
                                        $timeout = (Get-Date).AddMinutes(5)
                                        do {
                                            Start-Sleep -Seconds 2
                                            $taskInfo = Get-ScheduledTask -TaskName $taskName
                                            if ($taskInfo.State -eq 'Ready') { break }
                                        } while ((Get-Date) -lt $timeout)
                                        
                                        $taskResult = Get-ScheduledTaskInfo -TaskName $taskName
                                        Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
                                        
                                        if ($taskResult.LastTaskResult -ne 0) {
                                            throw "Task failed for user $userName with result: $($taskResult.LastTaskResult)"
                                        }
                                    }
                                    catch {
                                        Write-Host "Failed to process user $userName : $_"
                                    }
                                }
                                
                                $results += @{
                                    Context = 'EveryUserExec'
                                    Status = 'Success'
                                }
                            }
                            catch {
                                $results += @{
                                    Context = 'EveryUserExec'
                                    Status = 'Failed'
                                    Error = $_.ToString()
                                }
                            }
                        }
        
                        return $results
                    } -ArgumentList $remotePath, $RemediationName

                    # Convert hashtable to PSObject for consistent CSV export
                    return [PSCustomObject]@{
                        ComputerName = $ComputerName
                        Status = 'Success'
                        Results = $results
                        Error = $null
                        Timestamp = Get-Date
                    }
                }
                catch {
                    # Convert hashtable to PSObject for consistent CSV export
                    return [PSCustomObject]@{
                        ComputerName = $ComputerName
                        Status = 'Failed'
                        Results = $null
                        Error = $_.Exception.Message
                        Timestamp = Get-Date
                    }
                }
                finally {
                    if ($session) {
                        Remove-PSSession -Session $session
                    }
                }
            }).AddArgument($computerName).AddArgument($script:RemediationPath).AddArgument($RemoteTempDir).AddArgument($RemediationName)

            $powerShell.RunspacePool = $runspacePool

            $runspaces += @{
                PowerShell = $powerShell
                Handle = $powerShell.BeginInvoke()
                ComputerName = $computerName
                StartTime = Get-Date
            }
        }

        # Process completed jobs and handle results
        $completed = @()
        $timeout = (Get-Date).AddMinutes(30)  # 30-minute total timeout

        do {
            foreach ($runspace in $runspaces | Where-Object { $_.Handle -and !$completed.Contains($_.Handle) }) {
                if ($runspace.Handle.IsCompleted) {
                    $result = $runspace.PowerShell.EndInvoke($runspace.Handle)
                    $resultQueue.Enqueue($result)
                    $completed += $runspace.Handle
                    $runspace.PowerShell.Dispose()
                }
                elseif ((Get-Date) -gt $timeout) {
                    Write-Log "Timeout occurred for computer: $($runspace.ComputerName)" -Level Warning
                    $resultQueue.Enqueue([PSCustomObject]@{
                        ComputerName = $runspace.ComputerName
                        Status = 'Failed'
                        Error = 'Operation timed out'
                        Timestamp = Get-Date
                    })
                    $completed += $runspace.Handle
                    $runspace.PowerShell.Stop()
                    $runspace.PowerShell.Dispose()
                }
            }

            Start-Sleep -Milliseconds 100
        } while ($completed.Count -lt $runspaces.Count)

        # Clean up
        $runspacePool.Close()
        $runspacePool.Dispose()

        # Convert queue to array and return results
        $results = @()
        while ($resultQueue.TryDequeue([ref]$result)) {
            $results += $result
        }

        return $results
    }
}

Process {
    try {
        Write-Log "Starting parallel remediation deployment: $RemediationName" -Level Info
        Write-Log "Using remediation root: $RemediationsRoot" -Level Info
        Test-RemediationStructure

        $usingFailedComputers = $false
        $computerListPath = if (Test-Path -Path (Join-Path $script:RemediationPath 'failed_computers.txt')) {
            Write-Log "Using failed_computers.txt" -Level Info
            $usingFailedComputers = $true
            Join-Path $script:RemediationPath 'failed_computers.txt'
        }
        else {
            Join-Path $script:RemediationPath 'computers.txt'
        }

        Write-Log "Using computer list: $computerListPath" -Level Info
        $computers = @(Get-Content $computerListPath -ErrorAction Stop)
        Write-Log "Found $($computers.Count) computers to process" -Level Info

        # Execute remediation in parallel
        $results = Invoke-ParallelRemediation -ComputerNames $computers
        $results | Export-Csv -Path $script:ResultsPath -NoTypeInformation -Force

        # Ensure we have valid results before processing
        Write-Log "Processing final results..." -Level Info
        
        if ($null -eq $results) {
            Write-Log "No results returned from parallel execution" -Level Error
            $results = @([PSCustomObject]@{
                ComputerName = $computers[0]
                Status = "Failed"
                Error = "No results returned from parallel execution"
                Timestamp = Get-Date
                Results = $null
            })
        }
        
        $successful = @($results | Where-Object { $_.Status -eq 'Success' })
        $failed = @($results | Where-Object { $_.Status -eq 'Failed' })

        if ($successful.Count -gt 0) {
            $successful.ComputerName | Set-Content -Path (Join-Path $script:RemediationPath 'succeeded_computers.txt') -Force
            if (-not $usingFailedComputers -and $failed.Count -eq 0) {
                Remove-Item -Path (Join-Path $script:RemediationPath 'computers.txt') -Force -ErrorAction SilentlyContinue
            }
        }

        if ($failed.Count -gt 0) {
            $failed.ComputerName | Set-Content -Path (Join-Path $script:RemediationPath 'failed_computers.txt') -Force
        }
        else {
            Remove-Item -Path (Join-Path $script:RemediationPath 'failed_computers.txt') -Force -ErrorAction SilentlyContinue
        }

        Write-Log @"
Deployment Summary:
Total Computers: $($computers.Count)
Successful: $($successful.Count)
Failed: $($failed.Count)
Results exported to: $($script:ResultsPath)
"@ -Level Info
    }
    catch {
        Write-Log "Critical error in main execution: $_" -Level Error
        throw
    }
}

End {
    if ($script:TranscriptStarted) {
        Stop-Transcript
    }
}
