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
    [switch]$UseParallel,
    
    [Parameter(Mandatory = $false)]
    [int]$MaxParallelJobs = 5,
    
    [Parameter(Mandatory = $false)]
    [int]$ParallelTimeout = 1800,  # 30 minutes default timeout

    [Parameter(Mandatory = $false)]
    [ValidateSet('CurrentUser', 'System', 'Admin', 'EveryUser')]
    [string]$ExecutionContext = 'CurrentUser'
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

    # Create synchronized hashtable for thread-safe logging
    $script:SyncHash = [hashtable]::Synchronized(@{})
    $script:SyncHash.LogMutex = New-Object System.Threading.Mutex

    function Write-Log {
        param([string]$Message, [ValidateSet('Info', 'Warning', 'Error')][string]$Level = 'Info')
        try {
            $script:SyncHash.LogMutex.WaitOne() | Out-Null
            $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            $logMessage = "[$timestamp] [$Level] $Message"
            switch ($Level) {
                'Info' { Write-Host $logMessage -ForegroundColor Green }
                'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
                'Error' { Write-Host $logMessage -ForegroundColor Red }
            }
        }
        finally {
            $script:SyncHash.LogMutex.ReleaseMutex()
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

        $foundScripts = @('Remediate.ps1') |
            Where-Object { Test-Path (Join-Path $script:RemediationPath $_) } |
            ForEach-Object { Write-Log "Found execution script: $_" -Level Info; $true } |
            Test-Any

        if (!$foundScripts) {
            throw "Remediate.ps1 script not found"
        }
    }

    function Test-Any { 
        process { return $true } 
        end { return $false } 
    }

    function Get-ActiveUser {
        param([string]$ComputerName)
        
        try {
            Write-Host "Getting active user for $ComputerName"
            $quserOutput = quser /server:$ComputerName 2>&1 | Out-String
            Write-Host "Quser output: $quserOutput"
            
            $lines = $quserOutput -split '\r?\n' | Where-Object { $_ -match '\S' }
            if ($lines.Count -gt 1) {
                $userLines = $lines | Select-Object -Skip 1
                foreach ($line in $userLines) {
                    if ($line -match 'Active') {
                        $fields = $line.Trim() -split '\s+'
                        $user = $fields[0].Trim()
                        Write-Host "Found active user: $user"
                        return $user
                    }
                }
            }
            throw "No active user found"
        }
        catch {
            Write-Warning "Error getting active user: $_"
            return $null
        }
    }

    function Get-ValidUsers {
        param([string]$ComputerName)
        
        try {
            $validUserPattern = '^(?!(Public|defaultuser0|Default|.*\.adw|.*\.adf|Administrator|USAF_Admin)$)'
            $users = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_UserProfile | 
                Where-Object { $_.LocalPath -match '\\Users\\' } |
                ForEach-Object { Split-Path $_.LocalPath -Leaf } |
                Where-Object { $_ -match $validUserPattern }
            return $users
        }
        catch {
            Write-Warning "Error getting valid users: $_"
            return @()
        }
    }

    function Invoke-Remediation {
        param(
            [Parameter(Mandatory)]
            [string]$ComputerName
        )
        
        try {
            Write-Log "Processing computer: $ComputerName" -Level Info
            $session = New-PSSession -ComputerName $ComputerName
            $remotePath = Join-Path $RemoteTempDir $RemediationName

            # Setup remote directory
            Invoke-Command -Session $session -ScriptBlock { 
                param($Path) 
                if (-not (Test-Path $Path)) { 
                    New-Item -Path $Path -ItemType Directory -Force 
                } 
            } -ArgumentList $remotePath

            # Copy files
            Get-ChildItem $script:RemediationPath | 
                Where-Object { $_.Name -notin @('computers.txt', 'Logs') } | 
                Copy-Item -Destination $remotePath -ToSession $session -Recurse -Force

            # Execute remediation based on context
            $results = Invoke-Command -Session $session -ScriptBlock {
                param($remotePath, $RemediationName, $ExecutionContext)
                
                $scriptPath = Join-Path $remotePath "Remediate.ps1"
                if (-not (Test-Path $scriptPath)) {
                    throw "Remediation script not found: $scriptPath"
                }

                function Create-ScheduledTask {
                    param(
                        [string]$TaskName,
                        [string]$UserName,
                        [string]$ScriptPath,
                        [string]$LogPath,
                        [bool]$RunAsHighest = $true,
                        [string]$Description = ""
                    )

                    $action = New-ScheduledTaskAction -Execute "PowerShell.exe" `
                        -Argument "-ExecutionPolicy Bypass -WindowStyle Normal -NoProfile -File `"$ScriptPath`" -RemediationName `"$RemediationName`" -LogPath `"$LogPath`" -RemediationDir `"$remotePath`""

                    $trigger = New-ScheduledTaskTrigger -Once -At (Get-Date)
                    $principal = New-ScheduledTaskPrincipal -UserId $UserName -LogonType Interactive -RunLevel $(if ($RunAsHighest) { 'Highest' } else { 'Limited' })
                    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 5)

                    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force
                    return $TaskName
                }

                function Wait-ForTask {
                    param([string]$TaskName, [int]$TimeoutSeconds = 300)
                    
                    $timeout = (Get-Date).AddSeconds($TimeoutSeconds)
                    do {
                        Start-Sleep -Seconds 2
                        $taskInfo = Get-ScheduledTask -TaskName $TaskName
                        Write-Host "Task state: $($taskInfo.State)"
                        if ($taskInfo.State -eq 'Ready') { break }
                    } while ((Get-Date) -lt $timeout)

                    $taskResult = Get-ScheduledTaskInfo -TaskName $TaskName
                    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
                    return $taskResult.LastTaskResult
                }

                try {
                    switch ($ExecutionContext) {
                        'System' {
                            $taskName = "Remediation_${RemediationName}_System"
                            $taskName = Create-ScheduledTask -TaskName $taskName -UserName "SYSTEM" -ScriptPath $scriptPath `
                                -LogPath "$remotePath\system_result.log" -Description "System context remediation"
                        }
                        'Admin' {
                            $taskName = "Remediation_${RemediationName}_Admin"
                            $taskName = Create-ScheduledTask -TaskName $taskName -UserName "Administrator" -ScriptPath $scriptPath `
                                -LogPath "$remotePath\admin_result.log" -Description "Admin context remediation"
                        }
                        'EveryUser' {
                            $validUserPattern = '^(?!(Public|defaultuser0|Default|.*\.adw|.*\.adf|Administrator|USAF_Admin)$)'
                            $users = Get-ChildItem 'C:\Users' | 
                                Where-Object { $_.PSIsContainer -and $_.Name -match $validUserPattern }
                            
                            foreach ($userProfile in $users) {
                                $userName = $userProfile.Name
                                Write-Host "Processing user: $userName"
                                
                                $taskName = "Remediation_${RemediationName}_$userName"
                                $taskName = Create-ScheduledTask -TaskName $taskName -UserName $userName -ScriptPath $scriptPath `
                                    -LogPath "$remotePath\${userName}_result.log" -Description "User context remediation for $userName"
                                
                                Start-ScheduledTask -TaskName $taskName
                                $result = Wait-ForTask -TaskName $taskName
                                
                                if ($result -ne 0) {
                                    Write-Warning "Task failed for user $userName with result: $result"
                                }
                            }
                        }
                        default { # CurrentUser
                            $quserOutput = quser 2>&1 | Out-String
                            Write-Host "Quser output: $quserOutput"
                            
                            $currentUser = $null
                            $lines = $quserOutput -split '\r?\n' | Where-Object { $_ -match '\S' }
                            if ($lines.Count -gt 1) {
                                $userLines = $lines | Select-Object -Skip 1
                                foreach ($line in $userLines) {
                                    if ($line -match 'Active') {
                                        $fields = $line.Trim() -split '\s+'
                                        $currentUser = $fields[0].Trim()
                                        Write-Host "Found active user: $currentUser"
                                        break
                                    }
                                }
                            }
                            
                            if (-not $currentUser) {
                                throw "No active user detected"
                            }
                            
                            $validUserPattern = '^(?!(Public|defaultuser0|Default|.*\.adw|.*\.adf|Administrator|USAF_Admin)$)'
                            if (-not ($currentUser -match $validUserPattern)) {
                                throw "User $currentUser is not valid for execution"
                            }

                            $taskName = "Remediation_${RemediationName}_$currentUser"
                            $taskName = Create-ScheduledTask -TaskName $taskName -UserName $currentUser -ScriptPath $scriptPath `
                                -LogPath "$remotePath\current_user_result.log" -Description "Current user context remediation"
                        }
                    }

                    # Start and wait for task
                    Start-ScheduledTask -TaskName $taskName
                    $result = Wait-ForTask -TaskName $taskName
                    Write-Host "Task completed with result: $result"

                    # Create and return result object
                    $resultObj = [PSCustomObject]@{
                        Context = $ExecutionContext
                        Status = $(if ($result -eq 0) { 'Success' } else { 'Failed' })
                        Error = $(if ($result -ne 0) { "Task failed with result: $result" } else { $null })
                    }

                    $jsonResult = $resultObj | ConvertTo-Json -Compress
                    Write-Host "Returning result: $jsonResult"
                    return $jsonResult

                }
                catch {
                    Write-Host "Error in remediation: $_"
                    Write-Host "Error details: $($_.Exception)"
                    Write-Host "Stack trace: $($_.ScriptStackTrace)"
                    
                    $errorResult = [PSCustomObject]@{
                        Context = $ExecutionContext
                        Status = 'Failed'
                        Error = 'No valid result found'
                    }
                }
            } catch {
                Write-Warning "Error parsing result: $_"
                $remediationResult = [PSCustomObject]@{
                    Context = $ExecutionContext
                    Status = 'Failed'
                    Error = "Error parsing result: $_"
                }
            }
            
            $finalResult = [PSCustomObject]@{
                ComputerName = $ComputerName
                Status = $remediationResult.Status
                Error = $remediationResult.Error
                Timestamp = Get-Date
                Details = $remediationResult
                Context = $ExecutionContext
            }
            
            Write-Host "Final result for $ComputerName : $($finalResult | ConvertTo-Json -Depth 10)"
            
            return $finalResult
        }
        catch {
            Write-Log "Error processing $ComputerName : $_" -Level Error
            return [PSCustomObject]@{
                ComputerName = $ComputerName
                Status = 'Failed'
                Error = $_.Exception.Message
                Timestamp = Get-Date
                Details = $null
                Context = $ExecutionContext
            }
        }
        finally {
            if ($session) {
                Remove-PSSession -Session $session
                Write-Log "Cleaned up PS Session for $ComputerName" -Level Info
            }
        }
    }
}

Process {
    try {
        Write-Log "Starting remediation deployment: $RemediationName" -Level Info
        Write-Log "Using remediation root: $RemediationsRoot" -Level Info
        Write-Log "Execution context: $ExecutionContext" -Level Info
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

        # Process computers
        $results = if ($UseParallel) {
            Write-Log "Using parallel execution with max $MaxParallelJobs concurrent jobs" -Level Info
            
            $jobTracker = [System.Collections.Hashtable]::Synchronized(@{})
            $jobs = @()
            $parallelResults = @()
            $processedCount = 0
            
            foreach ($computer in $computers) {
                while ((Get-Job -State Running).Count -ge $MaxParallelJobs) {
                    $completed = Get-Job -State Completed
                    foreach ($job in $completed) {
                        $jobResult = Receive-Job -Job $job
                        $parallelResults += $jobResult
                        $processedCount++
                        Write-Log "Completed $processedCount of $($computers.Count) computers" -Level Info
                        Remove-Job -Job $job
                    }
                    Start-Sleep -Seconds 1
                }
                
                $jobParams = @{
                    ScriptBlock = {
                        param($Computer, $RemediationName, $RemediationsRoot, $RemoteTempDir, $ConnectionTimeout, $CleanupOnCompletion, $Credential, $ExecutionContext)
                        
                        ${function:Write-Log} = $using:function:Write-Log
                        ${function:Test-Any} = $using:function:Test-Any
                        ${function:Test-RemediationStructure} = $using:function:Test-RemediationStructure
                        ${function:Invoke-Remediation} = $using:function:Invoke-Remediation
                        
                        Invoke-Remediation -ComputerName $Computer
                    }
                    ArgumentList = @($computer, $RemediationName, $RemediationsRoot, $RemoteTempDir, $ConnectionTimeout, $CleanupOnCompletion, $Credential, $ExecutionContext)
                }
                
                $job = Start-Job @jobParams
                $jobs += $job
                $jobTracker[$job.Id] = $computer
                Write-Log "Started job for computer: $computer" -Level Info
            }
            
            Write-Log "Waiting for remaining jobs to complete..." -Level Info
            $remainingJobs = $jobs | Where-Object { $_.State -eq 'Running' }
            if ($remainingJobs) {
                Write-Log "Waiting up to $ParallelTimeout seconds for jobs to complete..." -Level Info
                Wait-Job -Job $remainingJobs -Timeout $ParallelTimeout | Out-Null
                
                foreach ($job in $remainingJobs) {
                    if ($job.State -eq 'Running') {
                        Write-Log "Job for $($jobTracker[$job.Id]) timed out after $ParallelTimeout seconds" -Level Warning
                        Stop-Job -Job $job
                        $parallelResults += [PSCustomObject]@{
                            ComputerName = $jobTracker[$job.Id]
                            Status = 'Failed'
                            Error = "Job timed out after $ParallelTimeout seconds"
                            Timestamp = Get-Date
                            Details = $null
                            Context = $ExecutionContext
                        }
                    } else {
                        $jobResult = Receive-Job -Job $job
                        $parallelResults += $jobResult
                    }
                    Remove-Job -Job $job
                }
            }
            
            $parallelResults
        } else {
            @($computers | ForEach-Object { Invoke-Remediation -ComputerName $_ })
        }

        Write-Log "Processing final results..." -Level Info
        Write-Host "Results before processing: $(ConvertTo-Json $results -Depth 10)"
        
        $successful = @($results | Where-Object { $_.Status -eq 'Success' })
        $failed = @($results | Where-Object { $_.Status -eq 'Failed' -or (-not $_.Status) })
        
        $results | Export-Csv -Path $script:ResultsPath -NoTypeInformation -Force
        
        if ($successful.Count -gt 0) {
            $successful.ComputerName | Set-Content -Path (Join-Path $script:RemediationPath 'succeeded_computers.txt') -Force
            
            if (-not $usingFailedComputers -and $failed.Count -eq 0) {
                Remove-Item -Path (Join-Path $script:RemediationPath 'computers.txt') -Force -ErrorAction SilentlyContinue
            }
        }

        if ($failed.Count -gt 0) {
            $failed.ComputerName | Set-Content -Path (Join-Path $script:RemediationPath 'failed_computers.txt') -Force
        } else {
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
    try {
        Write-Log "Starting cleanup..." -Level Info
        
        # Stop transcript if it was started
        if ($script:TranscriptStarted) {
            Stop-Transcript
            $script:TranscriptStarted = $false
        }
        
        # Clean up any leftover jobs
        $remainingJobs = Get-Job
        if ($remainingJobs) {
            Write-Log "Cleaning up $($remainingJobs.Count) remaining jobs..." -Level Info
            $remainingJobs | Remove-Job -Force -ErrorAction SilentlyContinue
        }
        
        # Clean up remote temp directories if specified
        if ($CleanupOnCompletion) {
            Write-Log "Cleaning up remote temp directories..." -Level Info
            foreach ($computer in $computers) {
                try {
                    $session = New-PSSession -ComputerName $computer -ErrorAction Stop
                    Invoke-Command -Session $session -ScriptBlock {
                        param($Path)
                        if (Test-Path $Path) {
                            Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
                        }
                    } -ArgumentList (Join-Path $RemoteTempDir $RemediationName)
                }
                catch {
                    Write-Log "Failed to clean up remote directory on $computer : $_" -Level Warning
                }
                finally {
                    if ($session) {
                        Remove-PSSession -Session $session
                    }
                }
            }
        }
        
        # Clean up mutex
        if ($script:SyncHash.LogMutex) {
            $script:SyncHash.LogMutex.Dispose()
        }
        
        Write-Log "Remediation deployment completed" -Level Info
    }
    catch {
        Write-Warning "Error during cleanup: $_"
    }
}$_.ToString()
                    }
                    return ($errorResult | ConvertTo-Json -Compress)
                }
            } -ArgumentList $remotePath, $RemediationName, $ExecutionContext

            Write-Host "Processing results for $ComputerName"
            Write-Host "Raw results: $results"
            
            try {
                if ($results -and $results.Trim().StartsWith('{')) {
                    $remediationResult = $results | ConvertFrom-Json
                    Write-Host "Parsed remediation result: $($remediationResult | ConvertTo-Json -Depth 10)"
                } else {
                    Write-Warning "No valid JSON result found"
                    $remediationResult = [PSCustomObject]@{
                        Context = $ExecutionContext
                        Status = 'Failed'
                        Error =
