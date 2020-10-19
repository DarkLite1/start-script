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
        
        $scriptParameters = $(
            ((Get-Command $scriptPathItem.FullName).Parameters.GetEnumerator() | Where-Object { $psBuildInParameters -notContains $_.Key }).Key)
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
        $userParameters = $ParameterPathItem | Get-Content -Raw | 
        ConvertFrom-Json
        Write-Verbose "User parameters '$userParameters'"
        #endregion

        #region Copy parameter file to log folder
        Write-Verbose "Copy input file to log folder '$($LogParams.LogFolder)'"
        Copy-Item -Path $ParameterPathItem.FullName -Destination "$LogFile - $($userParameters.ScriptName) - $($ParameterPathItem.Name)"
        #endregion
    
        #region Test only known script parameters are used
        $userParameters.PSObject.Properties.Name | Where-Object {
            $scriptParameters -notContains $_
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
        
        foreach ($p in $scriptParameters) {
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

        $errorMessage = "Invalid parameter file '$($ParameterPathItem.FullName)`r`n'`r`ERROR: $_`r`n`r`nScript parameters:`r`n$($scriptParameters -join `"`r`n`")" 

        $errorMessage | Out-File  "$LogFile - $($userParameters.ScriptName) - $($ParameterPathItem.Name) - ERROR.txt" -Encoding utf8 -Force -EA Ignore

        Send-MailHC -To $ScriptAdmin -Subject FAILURE -Priority High -Message $errorMessage -Header $ScriptName

        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $errorMessage"
        Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}