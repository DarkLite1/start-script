#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog

<#
    .SYNOPSIS
        Execute a script with its parameters coming from a .json file.

    .DESCRIPTION
        This script is designed to be triggered by a scheduled task. It will
        execute a single script with a single .json input file containing its
        named parameters.

        The parameter ScriptName is mandatory in the parameter file.

        A copy of the parameter file is always saved in the log folder.

        Whenever the child script fails due to a throw. a parameter validation
        error, a missing mandatory parameter, incorrect input, ... an error
        file and mail is generated to inform the admin and a copy of the error
        file and import file are saved in the log folder.

        When the child script fails, this script will exit with exit code 1.
        This will set the scheduled task to status to 'failed' and allows it
        to be restarted when needed.

        Using 'Exit 1' in a child script will not trigger failure to the task
        scheduler or a failed script at all. Please use 'throw' in your scripts
        instead.

    .PARAMETER ScriptPath
        Path to the script file that needs to be executed.

    .PARAMETER ParameterPath
        Path to the .json file that contains the named script parameters and
        their values.
 #>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptPath,
    [Parameter(Mandatory)]
    [String]$ParameterPath,
    [String]$ScriptName = 'Start script (All)',
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\Start script",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        $null = Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Test ScriptPath
        try {
            $scriptPathItem = Get-Item -Path $ScriptPath -EA Stop
        }
        catch {
            throw "Script file '$ScriptPath' not found"
        }

        if ($scriptPathItem.Extension -ne '.ps1') {
            throw "*Script file '$scriptPathItem' needs to have the extension '.ps1'"
        }
        #endregion

        #region Test ParameterPath
        try {
            $ParameterPathItem = Get-Item -Path $ParameterPath -EA Stop
        }
        catch {
            throw "Parameter file '$ParameterPath' not found"
        }

        if ($ParameterPathItem.Extension -ne '.json') {
            throw "*Parameter file '$ParameterPathItem' needs to have the extension '.json'"
        }
        #endregion

        #region Get script parameters
        $defaultValues = Get-DefaultParameterValuesHC -Path $scriptPathItem.FullName

        $psBuildInParameters = [System.Management.Automation.PSCmdlet]::CommonParameters
        $psBuildInParameters += [System.Management.Automation.PSCmdlet]::OptionalCommonParameters

        $scriptParameters = (Get-Command $scriptPathItem.FullName).Parameters.GetEnumerator() |
        Where-Object { $psBuildInParameters -notContains $_.Key }

        $userInfoList = foreach ($p in $scriptParameters.GetEnumerator()) {
            'Name: {0} Type: {1} Mandatory: {2} ' -f
            $p.Value.Name, $p.Value.ParameterType, $p.Value.Attributes.Mandatory
        }
        #endregion

        #region Load modules for child script
        $LoadModules = {
            Get-ChildItem ($env:PSModulePath -split ';') -EA Ignore |
            Where-Object Name -Like 'Toolbox*' | Import-Module
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

Process {
    Try {
        $job = $null

        #region Test valid .json file
        try {
            $userParameters = $ParameterPathItem | Get-Content -Raw |
            ConvertFrom-Json
            Write-Verbose "User parameters '$userParameters'"
        }
        catch {
            throw "Invalid parameter file: $_"
        }
        #endregion

        #region Copy parameter file to log folder
        Write-Verbose "Copy input file to log folder '$($LogParams.LogFolder)'"
        Copy-Item -Path $ParameterPathItem.FullName -Destination "$LogFile - $($userParameters.ScriptName) - $($ParameterPathItem.Name)"
        #endregion

        #region Test only valid script parameters are used
        $userParameters.PSObject.Properties.Name | Where-Object {
            $scriptParameters.Key -notContains $_
        } | ForEach-Object {
            $invalidParameter = $_
            throw "The parameter '$invalidParameter' is not accepted by script '$($scriptPathItem.FullName)'."
        }
        #endregion

        #region Test ScriptName is mandatory
        if (-not $userParameters.ScriptName) {
            throw "Parameter 'ScriptName' is mandatory"
        }
        #endregion

        #region Build argument list for Start-Job
        $startJobArgumentList = @()

        foreach ($p in $scriptParameters.GetEnumerator()) {
            $value = ''
            if ($defaultValues[$p.Key]) {
                $value = $defaultValues[$p.Key]
            }
            if ($userParameters.($p.Key)) {
                $value = $userParameters.($p.Key)
            }
            $value = foreach ($v in $value) {
                if ($v -is [String]) {
                    # convert '$env:' variables to strings
                    $ExecutionContext.InvokeCommand.ExpandString($v)
                }
                elseif (
                     ($v -is [PSCustomObject]) -and
                     ($p.Value.ParameterType.Name -eq 'HashTable')
                ) {
                    # when the script accepts type HashTable
                    # convert 'PSCustomObject' to hash table
                    $hashTable = @{}
                    $v.PSObject.properties | ForEach-Object {
                        $hashTable[$_.Name] = $_.Value
                    }
                    $hashTable
                }
                else { $v }
            }
            Write-Verbose "Parameter name '$($p.Key)' value '$value'"
            $startJobArgumentList += , $value
        }
        #endregion

        #region Start job
        Write-Verbose 'Start job'

        Write-EventLog @EventOutParams -Message (
            "Launch script:`n" +
            "`n- Name:`t`t" + $userParameters.ScriptName +
            "`n- Script:`t" + $scriptPathItem.FullName +
            "`n- ArgumentList:`t" + $startJobArgumentList)
        $StartJobParams = @{
            Name         = $userParameters.ScriptName
            # Running startup script threw an error: Unable: Unable to load one or more of the requested types. Retrieve the LoaderExceptions property for more information..
            # InitializationScript = $LoadModules
            LiteralPath  = $scriptPathItem.FullName
            ArgumentList = $startJobArgumentList
        }
        $job = Start-Job @StartJobParams
        #endregion

        #region Wait for job launch
        Write-Verbose 'Wait 5 seconds for initial job launch'
        Start-Sleep -Seconds 5
        #endregion

        #region Missing mandatory parameters set the state to 'Blocked'
        if ($job.State -eq 'Blocked') {
            Write-Verbose "Job '$($job.Name)' status '$($job.State)'"
            throw "Job status 'Blocked', have you provided all mandatory parameters?"
        }
        #endregion

        #region Wait for job to finish
        Write-Verbose 'Wait for job to finish'
        $null = $job | Wait-Job
        #endregion

        #region Return job results or fail on error
        Write-Verbose "Job '$($job.Name)' status '$($job.State)'"

        if ($job.State -eq 'Failed') {
            throw $job.ChildJobs[0].JobStateInfo.Reason.Message
        }
        else {
            $job | Receive-Job
        }
        #endregion

        Write-Verbose 'Script finished successfully'
    }
    Catch {
        Write-Warning $_

        #region Create error file
        $errorFileMessage = [ordered]@{
            errorMessage         = $_.Exception.Message
            scriptName           = $userParameters.ScriptName
            scriptParameters     = $userInfoList
            startJobArgumentList = $startJobArgumentList
            ScriptFile           = $ScriptPath
            ParameterFile        = $ParameterPath
        } | ConvertTo-Json -Depth 5 | Format-JsonHC

        $logFileFullName = "$LogFile - $($userParameters.ScriptName) - $($ParameterPathItem.BaseName) - ERROR.json"

        $errorFileMessage | Out-File $logFileFullName -Encoding utf8 -Force -EA Ignore
        #endregion

        #region Send mail
        $mailParams = @{
            To          = $ScriptAdmin
            Subject     = "FAILURE{0}" -f $(
                if ($userParameters.ScriptName) {
                    ' - ' + $userParameters.ScriptName
                })
            Priority    = 'High'
            Message     = "Script '<b>$($userParameters.ScriptName)</b>' failed with error:
                 <p>$_</p>
                 <p><i>* Check the attachment for details</i></p>"
            Header      = $ScriptName
            Attachments = $logFileFullName
        }
        Send-MailHC @mailParams
        #endregion

        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $errorFileMessage"

        #region Return failure to the task scheduler
        Exit 1
        #endregion
    }
    Finally {
        Get-Job | Remove-Job -Force
        Write-EventLog @EventEndParams
    }
}