#Requires -RunAsAdministrator
#Requires -Version 5.0

[CmdletBinding()]
param (
  [Parameter(Mandatory = $true)]
  [string]$RemediationName,

  [Parameter(Mandatory = $false)]
  [string]$RemediationsRoot = '\\Dm-usafe-01\86cs_r$\05_SCOO\01. SCOO (Applies to ALL of SCOO)\scripts\RemediationLauncher\Remediations',

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
  if (-not $Credential) {
    $currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    Write-Host "Running as user: $($currentUser.Name)"
  }

  if (-not (Test-Path $RemediationsRoot)) {
    throw "Cannot access remediation root path: $RemediationsRoot. Please verify network connectivity and permissions."
  }

  $remediationPath = Join-Path $RemediationsRoot $RemediationName
  $logFolder = Join-Path $remediationPath 'Logs'

  # Script settings
  $script:UseParallelExecution = $false  # Toggle this to enable/disable parallel processing
  $script:MaxConcurrentJobs = 10        # Maximum number of concurrent jobs when parallel execution is enabled
  $script:DateStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
  $script:MainLogPath = Join-Path $logFolder "Deployment_$DateStamp.log"
  $script:ResultsPath = Join-Path $logFolder "Results_$DateStamp.csv"
  $script:ErrorLog = Join-Path $logFolder "Errors_$DateStamp.log"

  if (!(Test-Path $logFolder)) { New-Item -ItemType Directory -Path $logFolder -Force | Out-Null }
  try { Start-Transcript -Path $script:MainLogPath -Force }
  catch { Write-Warning "Could not start transcript: $($_.Exception.Message)" }

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
    if (!(Test-Path $remediationPath)) { throw "Remediation folder not found: $remediationPath" }

    $computerListPath = Join-Path $remediationPath 'computers.txt'
    $failedComputerListPath = Join-Path $remediationPath 'failed_computers.txt'
    if (!(Test-Path $computerListPath) -and !(Test-Path $failedComputerListPath)) {
      throw 'computers.txt not found in remediation folder'
    }

    $foundScripts = @('AdminExec.ps1', 'CurrentUserExec.ps1', 'EveryUserExec.ps1', 'SystemExec.ps1') |
    Where-Object { Test-Path (Join-Path $remediationPath $_) } |
    ForEach-Object { Write-Log "Found execution script: $_" -Level Info; $true } |
    Test-Any

    if (!$foundScripts) { throw "No execution scripts found" }
  }

  # Helper function for PowerShell 5.1 compatibility (replacing .Any() LINQ method)
  function Test-Any { process { return $true } end { return $false } }
}

Process {
  function New-RemediationTask {
    param($Name, $Script, $User, $Arguments = @{}, [switch]$Start, [switch]$Remove)
    $action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument ("-ExecutionPolicy Bypass -WindowStyle Hidden -File `"$Script`" " +
            ($Arguments.GetEnumerator() | ForEach-Object { "-$($_.Key) `"$($_.Value)`"" } -join ' '))
    $trigger = if ($User -eq 'SYSTEM') { New-ScheduledTaskTrigger -Once -At (Get-Date) } else { New-ScheduledTaskTrigger -Once -AtLogOn -User $User }
    $principal = New-ScheduledTaskPrincipal -UserId $User -LogonType $(if ($User -eq 'SYSTEM') { 'ServiceAccount' } else { 'Interactive' }) -RunLevel Highest
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -Hidden
    Register-ScheduledTask -TaskName $Name -Action $action -Trigger $trigger -Principal $principal -Settings $settings -Force
    if ($Start) { Start-ScheduledTask -TaskName $Name }
    if ($Remove) { Unregister-ScheduledTask -TaskName $Name -Confirm:$false }
  }

  function Invoke-Remediation {
    param([Parameter(Mandatory)]$ComputerName)
    try {
      Write-Log "Processing computer: $ComputerName" -Level Info
      $session = New-PSSession @{ ComputerName = $ComputerName; ErrorAction = 'Stop' } + $(if ($Credential) { @{Credential = $Credential } } else { @{} })
      $remotePath = Join-Path $RemoteTempDir $RemediationName

      Invoke-Command -Session $session -ScriptBlock { param($Path) if (-not (Test-Path $Path)) { New-Item -Path $Path -ItemType Directory -Force } } -ArgumentList $remotePath
      Get-ChildItem $remediationPath | Where-Object { $_.Name -notin @('computers.txt', 'Logs') } | Copy-Item -Destination $remotePath -ToSession $session -Recurse -Force

      $results = Invoke-Command -Session $session -ScriptBlock {
        param($remotePath, $RemediationName)
        $results = @()
        $validUser = '^(?!(Public|defaultuser0|Default|.*\.adw|.*\.adf|Administrator|USAF_Admin)$)'

        @{
          'AdminExec.ps1'       = { param($script) & $script }
          'SystemExec.ps1'      = { param($script)
            New-RemediationTask -Name "Remediation_${RemediationName}_System" -Script $script -User 'SYSTEM' -Start -Remove -Args @{
              RemediationName = $RemediationName
              LogPath         = "$remotePath\result.log"
              RemediationDir  = $remotePath
            }
          }
          'EveryUserExec.ps1'   = { param($script)
            Get-ChildItem 'C:\Users' | Where-Object { $_.PSIsContainer -and $_.Name -match $validUser } |
            ForEach-Object { New-RemediationTask -Name "Remediation_${RemediationName}_$($_.Name)" -Script $script -User $_.Name }
          }
          'CurrentUserExec.ps1' = { param($script)
            $user = (Get-CimInstance Win32_ComputerSystem).UserName -split '\\' | Select-Object -Last 1
            if ($user -and $user -match $validUser -and (Test-Path "C:\Users\$user")) {
              New-RemediationTask -Name "Remediation_${RemediationName}_$user" -Script $script -User $user -Start
            }
          }
        }.GetEnumerator() | ForEach-Object {
          $script = Join-Path $remotePath $_.Key
          if (Test-Path $script) {
            try {
              & $_.Value $script
              $results += @{ Context = $_.Key.Replace('.ps1', ''); Status = 'Success'; Timestamp = Get-Date }
            }
            catch {
              $results += @{ Context = $_.Key.Replace('.ps1', ''); Status = 'Failed'; Error = $_.Exception.Message; Timestamp = Get-Date }
            }
          }
        }

        $results | ConvertTo-Json | Set-Content -Path (Join-Path $remotePath 'deployment_complete.flag')
        $results
      } -ArgumentList $remotePath, $RemediationName

      Write-Log "Remediation completed on $ComputerName" -Level Info
      $results
    }
    catch {
      Write-Log "Error processing $ComputerName : $_" -Level Error
      @{ ComputerName = $ComputerName; Status = 'Failed'; Error = $_.Exception.Message; Timestamp = Get-Date }
    }
    finally {
      if ($session) {
        Remove-PSSession -Session $session
        Write-Log "Cleaned up PS Session for $ComputerName" -Level Info
      }
    }
  }

  try {
    Write-Log "Starting remediation deployment: $RemediationName" -Level Info
    Write-Log "Using remediation root: $RemediationsRoot" -Level Info
    Test-RemediationStructure

    $computerListPath = if (Test-Path -Path (Join-Path $remediationPath 'failed_computers.txt')) {
      Write-Log "failed_computers.txt found... ignoring computers.txt" -Level Info
      Join-Path $remediationPath 'failed_computers.txt'
    }
    else {
      Join-Path $remediationPath 'computers.txt'
    }

    $computers = Get-Content $computerListPath -ErrorAction Stop
    Write-Log "Found $($computers.Count) computers to process" -Level Info

    if ($script:UseParallelExecution) {
      # Set up synchronized collections for parallel processing
      $script:SyncResults = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
      $script:SyncHash = [hashtable]::Synchronized(@{})

      # Create runspace pool
      $runspacePool = [runspacefactory]::CreateRunspacePool(1, $script:MaxConcurrentJobs)
      $runspacePool.Open()

      $jobs = @()

      foreach ($computer in $computers) {
        $powerShell = [powershell]::Create().AddScript({
            param($Computer, $Using:RemediationName, $Using:RemoteTempDir, $Using:Credential)

            # Import required functions into this runspace
            ${function:Write-Log} = $Using:WriteLogFunction
            ${function:New-RemediationTask} = $Using:NewRemediationTaskFunction
            ${function:Invoke-Remediation} = $Using:InvokeRemediationFunction

            $result = Invoke-Remediation -ComputerName $Computer
            $script:SyncResults.Add($result)
          }).AddArgument($computer).AddArgument($RemediationName).AddArgument($RemoteTempDir).AddArgument($Credential)

        $powerShell.RunspacePool = $runspacePool

        $jobs += @{
          PowerShell = $powerShell
          Handle     = $powerShell.BeginInvoke()
        }
      }

      # Monitor job progress
      do {
        $running = @($jobs | Where-Object { -not $_.Handle.IsCompleted }).Count
        if ($running -gt 0) {
          Write-Log "Processing $running computers..." -Level Info
          Start-Sleep -Seconds 10
        }
      } while ($running -gt 0)

      # Clean up jobs
      foreach ($job in $jobs) {
        $job.PowerShell.EndInvoke($job.Handle)
        $job.PowerShell.Dispose()
      }

      $runspacePool.Close()
      $runspacePool.Dispose()

      $results = @($script:SyncResults)
    }
    else {
      # Sequential processing
      $results = $computers | ForEach-Object { Invoke-Remediation -ComputerName $_ }
    }

    $results | Export-Csv -Path $script:ResultsPath -Force

    $successful = @($results | Where-Object Status -eq 'Success')
    $failed = @($results | Where-Object Status -eq 'Failed')

    if (-not $usingFailedComputers) { Remove-Item -Path $computerListPath -Force }

    $paths = @{
      Successful = Join-Path $remediationPath 'succeeded_computers.txt'
      Failed     = Join-Path $remediationPath 'failed_computers.txt'
    }

    foreach ($key in $paths.Keys) {
      $computers = $(if ($key -eq 'Successful') { $successful } else { $failed })
      if ($computers.Count) {
        Set-Content -Path $paths[$key] -Value $computers.ComputerName -Force
        Write-Log "Updated $($paths[$key])" -Level Info
      }
      elseif (Test-Path $paths[$key]) {
        Remove-Item -Path $paths[$key] -Force
        Write-Log "Deleted $($paths[$key])" -Level Info
      }
    }

    Write-Log @"
Deployment Summary:
Total Computers: $($computers.Count)
Successful: $($successful.Count)
Failed: $($failed.Count)
Results exported to: $script:ResultsPath
"@ -Level Info
  }
  catch {
    Write-Log "Critical error in main execution: $_" -Level Error
    throw
  }
}

End {
  Stop-Transcript
}
