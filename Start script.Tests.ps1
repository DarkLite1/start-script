#Requires -Module Assert, Pester

BeforeAll {
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $startInvokeCommand = Get-Command Invoke-Command

    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and 
        ($Priority -eq 'High') -and 
        ($Subject -eq 'FAILURE')
    }
    
    $testScriptPath = (New-Item -Path "TestDrive:\script.ps1" -Force -ItemType File -EA Ignore).FullName
    @"
    Param (
        [Parameter(Mandatory)]
        [String]`$PrinterName,
        [Parameter(Mandatory)]
        [String]`$PrinterColor,
        [String]`$ScriptName,
        [String]`$Tasks,
        [String]`$PaperSize = 'A4'
    )
"@ | Out-File -FilePath $testScriptPath -Encoding utf8 -Force

    $testParameterPath = (New-Item -Path "TestDrive:\inputFile.json" -Force -ItemType File -EA Ignore).FullName

    @{  
        PrinterColor = 'red'
        PrinterName  = "MyCustomPrinter"
        ScriptName   = 'Get printers'
    } | ConvertTo-Json | Out-File $testParameterPath -Encoding utf8

    $Params = @{
        ScriptPath    = $testScriptPath
        ParameterPath = $testParameterPath
        ScriptName    = 'Test'
        LogFolder     = (New-Item -Path "TestDrive:\Log" -ItemType Directory -EA Ignore).FullName
    }    

    Mock Invoke-Command
    Mock Send-MailHC
    Mock Write-EventLog
}

Describe 'error handling' {    
    Context 'mandatory parameters' {
        It '<Name>' -TestCases @(
            @{  Name = 'ScriptPath' }
            @{  Name = 'ParameterPath' }
        ) {
            (Get-Command $testScript).Parameters[$Name].Attributes.Mandatory |
            Should -BeTrue
        }
    }
    Context 'the logFolder' {
        It 'should exist' {
            $clonedParams = $Params.Clone()
            $clonedParams.LogFolder = 'NotExistingLogFolder'
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Path 'NotExistingLogFolder' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
    }
    Context 'the ScriptPath' {
        It 'should exist' {
            $clonedParams = $Params.Clone()
            $clonedParams.ScriptPath = 'NotExisting'
            
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Script file 'NotExisting' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
        It 'should have the extension .ps1' {
            $clonedParams = $Params.Clone()
            $clonedParams.ScriptPath = (New-Item -Path "TestDrive:\incorrectScript.txt" -Force -ItemType File -EA Ignore).FullName
            
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Script file '*incorrectScript.txt' needs to have the extension '.ps1'*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
    }
    Context 'the ParameterPath' {
        It 'should exist' {
            $clonedParams = $Params.Clone()
            $clonedParams.ParameterPath = 'NotExisting'
            
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Parameter file 'NotExisting' not found*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
        It 'should have the extension .json' {
            $clonedParams = $Params.Clone()
            $clonedParams.ParameterPath = (New-Item -Path "TestDrive:\incorrectParameter.txt" -Force -ItemType File -EA Ignore).FullName
            
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Parameter file '*incorrectParameter.txt' needs to have the extension '.json'*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter { $EntryType -eq 'Error' }
        }
    }
}

Describe 'a valid parameter input file' {
    BeforeAll {
        . $testScript @Params
    }
    It 'should retrieve the parameters from the .json parameter file' {
        $userParameters | Should -Not -BeNullOrEmpty
    }
    It 'should have the parameters ScriptName in the .json parameter file' {
        $userParameters.ScriptName | Should -Not -BeNullOrEmpty
    }
    It 'should invoke Invoke-Command with the parameters in the correct order' {
        Should -invoke Invoke-Command -Exactly 1 -Scope Describe -ParameterFilter {
            ($ErrorAction -eq 'Stop') -and
            ($FilePath -eq $testScriptPath) -and
            ($ComputerName -eq $env:COMPUTERNAME) -and
            ($ArgumentList[0] -eq 'MyCustomPrinter') -and
            ($ArgumentList[1] -eq 'red') -and
            ($ArgumentList[2] -eq 'Get printers') -and
            ($ArgumentList[3] -eq $null) -and
            ($ArgumentList[4] -eq 'A4') # default parameter in the script is copied
        }
    }
    Context 'logging' {
        BeforeAll {
            $testLogFolder = "$($Params.LogFolder)\Start script"
            $testLogFile = Get-ChildItem $testLogFolder -Recurse -File
        }
        It 'the log folder is created' {            
            $testLogFolder | Should -Exist
        }
        It 'the parameter input file is copied to the log folder' {
            $testLogFile.Count | Should -BeExactly 1
            $testLogFile.Name | Should -BeLike '*- Get printers - inputFile.json'
        }
    }
}

Describe 'when the parameter file is not valid because' {
    Context 'it is missing the property ScriptName' {
        BeforeAll {
            $testInputFile = (New-Item -Path "TestDrive:\InputFile.json" -Force -ItemType File -EA Ignore).FullName
     
            @{
                ScriptName   = $null
                PrinterColor = 'red'
                PrinterName  = "MyCustomPrinter"
            } | ConvertTo-Json | Out-File $testInputFile -Encoding utf8

            $clonedParams = $Params.Clone()
            $clonedParams.ParameterPath = $testInputFile

            . $testScript @clonedParams
        }
        It 'Invoke-Command is not called' {
            Should -Not -Invoke Invoke-Command -Scope Context
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Parameter 'ScriptName' is mandatory*")
            }
        }
        Context 'logging' {
            BeforeAll {
                $testLogFolder = "$($Params.LogFolder)\Start script"
                $testLogFile = Get-ChildItem $testLogFolder -Recurse -File
            }
            It 'the log folder is created' {            
                $testLogFolder | Should -Exist
            }
            It 'two files are created in the log folder' {
                $testLogFile.Count | Should -BeExactly 2
            }
            It 'one file is a copy of the input file' {
                $testLogFile[0].Name | Should -BeLike '*-  - inputFile.json'
            }
            It 'the other file contains the error message' {
                $testLogFile[1].Name | Should -BeLike '*-  - inputFile.json - ERROR.txt'
            }
        }
    }
    Context 'it is not a valid .json file' {
        BeforeAll {
            $testInputFile = (New-Item -Path "TestDrive:\InputFile.json" -Force -ItemType File -EA Ignore).FullName
     
            "NotJsonFormat ;!= " | Out-File $testInputFile -Encoding utf8

            $clonedParams = $Params.Clone()
            $clonedParams.ParameterPath = $testInputFile

            . $testScript @clonedParams
        }
        It 'Invoke-Command is not called' {
            Should -Not -Invoke Invoke-Command -Scope Context
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Invalid parameter file*")
            }
        }
        Context 'logging' {
            BeforeAll {
                $testLogFolder = "$($Params.LogFolder)\Start script"
                $testLogFile = Get-ChildItem $testLogFolder -Recurse -File
            }
            It 'the log folder is created' {            
                $testLogFolder | Should -Exist
            }
            It 'two files are created in the log folder' {
                $testLogFile.Count | Should -BeExactly 1
            }
            It 'the other file contains the error message' {
                $testLogFile[0].Name | Should -BeLike '*-  - inputFile.json - ERROR.txt'
            }
        }
    }
    Context 'the user used a parameter that is not available in the script' {
        BeforeAll {
            $testInputFile = (New-Item -Path "TestDrive:\InputFile.json" -Force -ItemType File -EA Ignore).FullName
     
            @{
                PrinterColor     = 'red'
                PrinterName      = "MyCustomPrinter"
                ScriptName       = 'Get printers'
                UnknownParameter = 'kiwiw'
            } | ConvertTo-Json | Out-File $testInputFile -Encoding utf8

            $clonedParams = $Params.Clone()
            $clonedParams.ParameterPath = $testInputFile

            . $testScript @clonedParams
        }
        It 'Invoke-Command is not called' {
            Should -Not -Invoke Invoke-Command -Scope Context
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*parameter 'UnknownParameter' is not accepted by script*")
            }
        }
        Context 'logging' {
            BeforeAll {
                $testLogFolder = "$($Params.LogFolder)\Start script"
                $testLogFile = Get-ChildItem $testLogFolder -Recurse -File
            }
            It 'the log folder is created' {            
                $testLogFolder | Should -Exist
            }
            It 'two files are created in the log folder' {
                $testLogFile.Count | Should -BeExactly 2
            }
            It 'one file is a copy of the input file' {
                $testLogFile[0].Name | Should -BeLike '*- Get printers - inputFile.json'
            }
            It 'the other file contains the error message' {
                $testLogFile[1].Name | Should -BeLike '*- Get printers - inputFile.json - ERROR.txt'
            }
        }
    }
    Context 'when Invoke-Command fails because of an incorrect parameter in the input file' {
        BeforeAll {
            $testInputFile = (New-Item -Path "TestDrive:\InputFile.json" -Force -ItemType File -EA Ignore).FullName
     
            @{  
                PrinterColor = 'red'
                PrinterName  = "MyCustomPrinter"
                ScriptName   = 'Get printers'
            } | ConvertTo-Json | Out-File $testParameterPath -Encoding utf8
    
            $clonedParams = $Params.Clone()
            $clonedParams.ParameterPath = $testInputFile

            . $testScript @clonedParams

            Mock Invoke-Command {
                & $startInvokeCommand -Scriptblock { 
                    Param (
                        [parameter(Mandatory)]
                        [int]$validParameter
                    )
                } -ArgumentList 'string'
            }
            . $testScript @Params
        }
        It 'Invoke-Command is called' {
            Should -Invoke Invoke-Command -Scope Context
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Input string was not in a correct format*")
            }
        }
        Context 'logging' {
            BeforeAll {
                $testLogFolder = "$($Params.LogFolder)\Start script"
                $testLogFile = Get-ChildItem $testLogFolder -Recurse -File
            }
            It 'the log folder is created' {            
                $testLogFolder | Should -Exist
            }
            It 'two files are created in the log folder' {
                $testLogFile.Count | Should -BeExactly 2
            }
            It 'one file is a copy of the input file' {
                $testLogFile[0].Name | Should -BeLike '*- Get printers - inputFile.json'
            }
            It 'the other file contains the error message' {
                $testLogFile[1].Name | Should -BeLike '*- Get printers - inputFile.json - ERROR.txt'
            }
        }
    } -Tag test
}
