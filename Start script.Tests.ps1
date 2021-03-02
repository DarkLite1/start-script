#Requires -Module Pester

BeforeAll {
    $StartJobCommand = Get-Command Start-Job

    $testParams = @{
        ScriptPath    = (New-Item -Path "TestDrive:\script.ps1" -Force -ItemType File).FullName
        ParameterPath = (New-Item -Path "TestDrive:\inputFile.json" -Force -ItemType File).FullName
        ScriptName    = 'Test'
        LogFolder     = New-Item -Path "TestDrive:\Log" -ItemType Directory
    }
    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')

    $MailAdminParams = {
        ($To -eq $ScriptAdmin) -and 
        ($Priority -eq 'High') -and 
        ($Subject -eq 'FAILURE')
    }
    
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
"@ | Out-File $testParams.ScriptPath -Encoding utf8 -Force

    @{  
        PrinterColor = 'red'
        PrinterName  = "MyCustomPrinter"
        ScriptName   = 'Get printers'
    } | ConvertTo-Json | Out-File $testParams.ParameterPath -Encoding utf8

    Mock Start-Job -MockWith {
        & $StartJobCommand -Scriptblock { 1 } -Name 'Get printers (BNL)'
    }
    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'error handling' {    
    Context 'mandatory parameters' {
        It '<_>' -TestCases @('ScriptPath' , 'ParameterPath' ) {
            (Get-Command $testScript).Parameters[$_].Attributes.Mandatory |
            Should -BeTrue
        }
    }
    Context 'the logFolder' {
        It 'should exist' {
            $clonedParams = $testParams.Clone()
            $clonedParams.LogFolder = 'NotExistingLogFolder'
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Path 'NotExistingLogFolder' not found*")
            }
        }
    }
    Context 'the ScriptPath' {
        It 'should exist' {
            $clonedParams = $testParams.Clone()
            $clonedParams.ScriptPath = 'NotExisting'
            
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Script file 'NotExisting' not found*")
            }
        }
        It 'should have the extension .ps1' {
            $clonedParams = $testParams.Clone()
            $clonedParams.ScriptPath = (New-Item -Path "TestDrive:\incorrectScript.txt" -Force -ItemType File -EA Ignore).FullName
            
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Script file '*incorrectScript.txt' needs to have the extension '.ps1'*")
            }
        }
    }
    Context 'the ParameterPath' {
        It 'should exist' {
            $clonedParams = $testParams.Clone()
            $clonedParams.ParameterPath = 'NotExisting'
            
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Parameter file 'NotExisting' not found*")
            }
        }
        It 'should have the extension .json' {
            $clonedParams = $testParams.Clone()
            $clonedParams.ParameterPath = (New-Item -Path "TestDrive:\incorrectParameter.txt" -Force -ItemType File -EA Ignore).FullName
            
            . $testScript @clonedParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Parameter file '*incorrectParameter.txt' needs to have the extension '.json'*")
            }
        }
    }
}
Describe 'a valid parameter input file' {
    BeforeAll {
        . $testScript @testParams
    }
    It 'should retrieve the parameters from the .json parameter file' {
        $userParameters | Should -Not -BeNullOrEmpty
    }
    It 'should have the parameters ScriptName in the .json parameter file' {
        $userParameters.ScriptName | Should -Not -BeNullOrEmpty
    }
    It 'should invoke Start-Job with the parameters in the correct order' {
        Should -Invoke Start-Job -Exactly 1 -Scope Describe -ParameterFilter {
            ($Name -eq 'Get printers') -and
            ($LiteralPath -eq $testParams.ScriptPath) -and
            ($ArgumentList[0] -eq 'MyCustomPrinter') -and
            ($ArgumentList[1] -eq 'red') -and
            ($ArgumentList[2] -eq 'Get printers') -and
            ($ArgumentList[3] -eq '') -and
            ($ArgumentList[4] -eq 'A4') # default parameter in the script
        }
    }
    Context 'logging' {
        BeforeAll {
            $testLogFolder = "$($testParams.LogFolder)\Start script"
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
Describe 'when Start-Job fails' {
    Context 'because of a missing mandatory parameter' {
        BeforeAll {
            Mock Start-Job {
                & $StartJobCommand -Scriptblock { 
                    Param (
                        [parameter(Mandatory)]
                        [int]$Number,
                        [parameter(Mandatory)]
                        [String]$Name
                    )
                    # parameters name is missing
                } -ArgumentList 1
            }
            . $testScript @testParams -EA SilentlyContinue
        }
        It 'Start-Job is called' {
            Should -Invoke Start-Job -Scope Context
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq $ScriptAdmin) -and 
                ($Priority -eq 'High') -and 
                ($Subject -eq 'FAILURE - Get printers') -and
                ($Message -like "*Job status 'Blocked', have you provided all mandatory parameters?*")
            }
        }
        Context 'logging' {
            BeforeAll {
                $testLogFolder = "$($testParams.LogFolder)\Start script"
                $testLogFile = Get-ChildItem $testLogFolder -Recurse -File
            }
            It 'the log folder is created' {            
                $testLogFolder | Should -Exist
            }
            It 'two files are created in the log folder' {
                $testLogFile.Count | Should -BeExactly 2
            }
            It 'one file is a copy of the input file' {
                $testLogFile[1].Name | Should -BeLike '*- Get printers - inputFile.json'
            }
            It 'the other file contains the error message' {
                $testLogFile[0].Name | Should -BeLike '*- Get printers - inputFile - ERROR.json'
                $actual = $testLogFile[0] | Get-Content -Raw | ConvertFrom-Json
                $actual.errorMessage | Should -BeLike "*Job status 'Blocked', have you provided all mandatory parameters?*"
            }
        }
    }
    Context 'because of parameter validation issues' {
        BeforeAll {
            Mock Start-Job {
                & $StartJobCommand -Scriptblock { 
                    Param (
                        [parameter(Mandatory)]
                        [int]$IncorrectParameters
                    )
                    # parameters not matching
                } -ArgumentList 'string'
            }
            . $testScript @testParams
        }
        It 'Start-Job is called' {
            Should -Invoke Start-Job -Scope Context
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq $ScriptAdmin) -and 
                ($Priority -eq 'High') -and 
                ($Subject -eq 'FAILURE - Get printers') -and
                ($Message -like "*Cannot process argument transformation on parameter 'IncorrectParameters'*")
            }
        }
        Context 'logging' {
            BeforeAll {
                $testLogFolder = "$($testParams.LogFolder)\Start script"
                $testLogFile = Get-ChildItem $testLogFolder -Recurse -File
            }
            It 'the log folder is created' {            
                $testLogFolder | Should -Exist
            }
            It 'two files are created in the log folder' {
                $testLogFile.Count | Should -BeExactly 2
            }
            It 'one file is a copy of the input file' {
                $testLogFile[1].Name | Should -BeLike '*- Get printers - inputFile.json'
            }
            It 'the other file contains the error message' {
                $testLogFile[0].Name | Should -BeLike '*- Get printers - inputFile - ERROR.json'
                $actual = $testLogFile[0] | Get-Content -Raw | ConvertFrom-Json
                $actual.errorMessage | Should -BeLike "*Cannot process argument transformation on parameter 'IncorrectParameters'*"
            }
        }
    }
    Context 'because of terminating errors in the execution script' {
        BeforeAll {
            Mock Start-Job {
                & $StartJobCommand -Scriptblock { 
                    throw 'Failure in script'
                }
            }
            . $testScript @testParams
        }
        It 'Start-Job is called' {
            Should -Invoke Start-Job -Scope Context
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq $ScriptAdmin) -and 
                ($Priority -eq 'High') -and 
                ($Subject -eq 'FAILURE - Get printers') -and
                ($Message -like "*Failure in script*")
            }
        }
        Context 'logging' {
            BeforeAll {
                $testLogFolder = "$($testParams.LogFolder)\Start script"
                $testLogFile = Get-ChildItem $testLogFolder -Recurse -File
            }
            It 'the log folder is created' {            
                $testLogFolder | Should -Exist
            }
            It 'two files are created in the log folder' {
                $testLogFile.Count | Should -BeExactly 2
            }
            It 'one file is a copy of the input file' {
                $testLogFile[1].Name | Should -BeLike '*- Get printers - inputFile.json'
            }
            It 'the other file contains the error message' {
                $testLogFile[0].Name | Should -BeLike '*- Get printers - inputFile - ERROR.json'
                $actual = $testLogFile[0] | Get-Content -Raw | ConvertFrom-Json
                $actual.errorMessage | Should -BeLike "*Failure in script*"
            }
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

            $clonedParams = $testParams.Clone()
            $clonedParams.ParameterPath = $testInputFile

            . $testScript @clonedParams
        }
        It 'Start-Job is not called' {
            Should -Not -Invoke Start-Job -Scope Context
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Parameter 'ScriptName' is mandatory*")
            }
        }
        Context 'logging' {
            BeforeAll {
                $testLogFolder = "$($testParams.LogFolder)\Start script"
                $testLogFile = Get-ChildItem $testLogFolder -Recurse -File
            }
            It 'the log folder is created' {            
                $testLogFolder | Should -Exist
            }
            It 'two files are created in the log folder' {
                $testLogFile.Count | Should -BeExactly 2
            }
            It 'one file is a copy of the input file' {
                $testLogFile[1].Name | Should -BeLike '*-  - inputFile.json'
            }
            It 'the other file contains the error message' {
                $testLogFile[0].Name | Should -BeLike '*-  - inputFile - ERROR.json'
                $actual = $testLogFile[0] | Get-Content -Raw | ConvertFrom-Json
                $actual.errorMessage | Should -BeLike "*Parameter 'ScriptName' is mandatory*"
            }
        }
    }
    Context 'it is not a valid .json file' {
        BeforeAll {
            $testInputFile = (New-Item -Path "TestDrive:\InputFile.json" -Force -ItemType File -EA Ignore).FullName
     
            "NotJsonFormat ;!= " | Out-File $testInputFile -Encoding utf8

            $clonedParams = $testParams.Clone()
            $clonedParams.ParameterPath = $testInputFile

            . $testScript @clonedParams
        }
        It 'Start-Job is not called' {
            Should -Not -Invoke Start-Job -Scope Context
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "*Invalid parameter file*")
            }
        }
        Context 'logging' {
            BeforeAll {
                $testLogFolder = "$($testParams.LogFolder)\Start script"
                $testLogFile = Get-ChildItem $testLogFolder -Recurse -File
            }
            It 'the log folder is created' {            
                $testLogFolder | Should -Exist
            }
            It 'two files are created in the log folder' {
                $testLogFile.Count | Should -BeExactly 1
            }
            It 'the other file contains the error message' {
                $testLogFile[0].Name | Should -BeLike '*-  - inputFile - ERROR.json'
                $actual = $testLogFile[0] | Get-Content -Raw | ConvertFrom-Json
                $actual.errorMessage | Should -BeLike "*Invalid parameter file*"
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
                UnknownParameter = 'kiwi'
            } | ConvertTo-Json | Out-File $testInputFile -Encoding utf8

            $clonedParams = $testParams.Clone()
            $clonedParams.ParameterPath = $testInputFile

            . $testScript @clonedParams
        }
        It 'Start-Job is not called' {
            Should -Not -Invoke Start-Job -Scope Context
        }
        It 'an email is sent to the admin' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Context -ParameterFilter {
                ($To -eq $ScriptAdmin) -and 
                ($Priority -eq 'High') -and 
                ($Subject -eq 'FAILURE - Get printers') -and
                ($Message -like "*parameter 'UnknownParameter' is not accepted by script*")
            }
        }
        Context 'logging' {
            BeforeAll {
                $testLogFolder = "$($testParams.LogFolder)\Start script"
                $testLogFile = Get-ChildItem $testLogFolder -Recurse -File
            }
            It 'the log folder is created' {            
                $testLogFolder | Should -Exist
            }
            It 'two files are created in the log folder' {
                $testLogFile.Count | Should -BeExactly 2
            }
            It 'one file is a copy of the input file' {
                $testLogFile[1].Name | Should -BeLike '*- Get printers - inputFile.json'
            }
            It 'the other file contains the error message' {
                $testLogFile[0].Name | Should -BeLike '*- Get printers - inputFile - ERROR.json'
                $actual = $testLogFile[0] | Get-Content -Raw | ConvertFrom-Json
                $actual.errorMessage | Should -BeLike "*parameter 'UnknownParameter' is not accepted by script*"
            }
        }
    }
}
Describe 'invoke Start-Job is called with the correct argument type when' {
    Context 'the arguments are in the input .JSON file' {
        BeforeAll {
            @"
        Param (
            [Parameter(Mandatory)]
            [String]`$ScriptName, 
            [Parameter(Mandatory)]
            [String[]]`$Colors,  
            [Parameter(Mandatory)]
            [PSCustomObject]`$CustomObject,
            [Parameter(Mandatory)]
            [HashTable]`$CustomHashTable,
            [String]`$LogFolder = "`$testLogFolder",
            [String]`$customLogFolder = "`$testLogFolder\`$ScriptName"
        )
"@ | Out-File $testParams.ScriptPath -Encoding utf8 -Force
    
            @{
                ScriptName      = 'Get printers'
                Colors          = @('red', 'green', 'blue')
                CustomObject    = @{
                    Duplex = 'Yes'
                }
                CustomHashTable = @{
                    DoubleSided = 'No'
                }
            } | ConvertTo-Json | 
            Out-File $testParams.ParameterPath -Encoding utf8  -Force
     
            . $testScript @testParams
        }
        It 'string' {
            Should -Invoke Start-Job -Exactly 1 -Scope Context -ParameterFilter {
                ($LiteralPath -eq $testParams.ScriptPath) -and
                ($ArgumentList[0] -eq 'Get printers') -and
                ($ArgumentList[0] -is [String])
            }
        }
        It 'array' {
            Should -Invoke Start-Job -Exactly 1 -Scope Context -ParameterFilter {
                ($LiteralPath -eq $testParams.ScriptPath) -and
                ($ArgumentList[1][0] -eq 'red') -and
                ($ArgumentList[1][1] -eq 'green') -and
                ($ArgumentList[1][2] -eq 'blue') -and
                ($ArgumentList[1] -is [System.Array])
            }
        }
        It 'PSCustomObject' {
            Should -Invoke Start-Job -Exactly 1 -Scope Context -ParameterFilter {
                ($LiteralPath -eq $testParams.ScriptPath) -and
                ($ArgumentList[2].Duplex -eq 'Yes') -and
                ($ArgumentList[2] -is [PSCustomObject])
            }
        }
        It 'HashTable when the script parameter is of type HashTable' {
            Should -Invoke Start-Job -Exactly 1 -Scope Context -ParameterFilter {
                ($LiteralPath -eq $testParams.ScriptPath) -and
                ($ArgumentList[3].DoubleSided -eq 'No') -and
                ($ArgumentList[3] -is [HashTable])
            }
        }
    }
    Context 'the arguments are default values in the script file' {
        BeforeAll {
            @"
            Param (
                [Parameter(Mandatory)]
                [String]`$ScriptName, 
                [String[]]`$Colors = @('red', 'green'),  
                [PSCustomObject]`$CustomObject = [PSCustomObject]@{
                    Duplex = 'Yes'
                },
                [HashTable]`$CustomHashTable = @{
                    DoubleSided = 'No'
                }
            )
"@ | Out-File $testParams.ScriptPath -Encoding utf8 -Force
        
            @{
                ScriptName = 'Get printers'
            } | ConvertTo-Json | 
            Out-File $testParams.ParameterPath -Encoding utf8  -Force
         
            . $testScript @testParams
        }
        It 'string' {
            Should -Invoke Start-Job -Exactly 1 -Scope Context -ParameterFilter {
                ($LiteralPath -eq $testParams.ScriptPath) -and
                ($ArgumentList[0] -eq 'Get printers') -and
                ($ArgumentList[0] -is [String])
            }
        }
        It 'array' {
            Should -Invoke Start-Job -Exactly 1 -Scope Context -ParameterFilter {
                ($LiteralPath -eq $testParams.ScriptPath) -and
                ($ArgumentList[1][0] -eq 'red') -and
                ($ArgumentList[1][1] -eq 'green') -and
                ($ArgumentList[1] -is [System.Array])
            }
        }
        It 'PSCustomObject' {
            Should -Invoke Start-Job -Exactly 1 -Scope Context -ParameterFilter {
                ($LiteralPath -eq $testParams.ScriptPath) -and
                ($ArgumentList[2].Duplex -eq 'Yes') -and
                ($ArgumentList[2] -is [PSCustomObject])
            }
        }
        It 'HashTable when the script parameter is of type HashTable' {
            Should -Invoke Start-Job -Exactly 1 -Scope Context -ParameterFilter {
                ($LiteralPath -eq $testParams.ScriptPath) -and
                ($ArgumentList[3].DoubleSided -eq 'No') -and
                ($ArgumentList[3] -is [HashTable])
            }
        } -Tag test
    }
}