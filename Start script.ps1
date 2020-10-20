<#
    .SYNOPSIS
        Execute a single script with parameters coming from a .json file.

    .DESCRIPTION
        This script is designed to be triggered by a scheduled task. It will
        execute a single script with a single .json input file containing its
        named parameters.

    .PARAMETER ScriptPath
        Path to the script that needs to be executed.

    .PARAMETER ParameterPath
        Path to the .json file that contains the script parameters.
 #>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptPath,
    [Parameter(Mandatory)]
    [String]$ParameterPath,
    [String]$ScriptName = 'Start script (All)', 
    [String]$LogFolder = "\\$env:COMPUTERNAME\Log",
    [String]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Try {
        $null = Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
        $LogParams = @{
            LogFolder    = New-FolderHC -Path $LogFolder -ChildPath 'Start script'
            Name         = $ScriptName
            Date         = 'ScriptStartTime'
            NoFormatting = $true
        }
        $LogFile = New-LogFileNameHC @LogParams
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

        $scriptParametersList = $($scriptParameters.Key)
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
            $scriptParametersList -notContains $_
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
    
        #region Build argument list for Invoke-Command
        $invokeCommandArgumentList = @()
        
        foreach ($p in $scriptParametersList) {
            $value = $null
            if ($defaultValues[$p]) {
                $value = $defaultValues[$p]
            }
            if ($userParameters.$p) {
                $value = $userParameters.$p
            }
            $value = $ExecutionContext.InvokeCommand.ExpandString($value)
            Write-Verbose "Parameter name '$p' value '$value'"
            $invokeCommandArgumentList += , $value
        }
        #endregion
        
        #region Start script with arguments in the correct order
        $invokeCommandParams = @{
            ErrorAction  = 'Stop'
            FilePath     = $scriptPathItem.FullName
            ComputerName = $env:COMPUTERNAME
            ArgumentList = $invokeCommandArgumentList 
        }
        Invoke-Command @invokeCommandParams
        #endregion
    }
    Catch {
        Write-Warning $_

        #region Create error file
        $errorFileMessage = [ordered]@{
            errorMessage              = $_.Exception.Message
            scriptName                = $userParameters.ScriptName
            scriptParameters          = $userInfoList
            invokeCommandArgumentList = $invokeCommandArgumentList
            ScriptFile                = $ScriptPath
            ParameterFile             = $ParameterPath
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
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}