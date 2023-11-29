#Requires -Modules Pester
#Requires -Version 5.1

BeforeDiscovery {
    # used by inModuleScope
    $testModule = $PSCommandPath.Replace('.Tests.ps1', '.psm1')
    $testModuleName = $testModule.Split('\')[-1].TrimEnd('.psm1')

    Remove-Module $testModuleName -Force -Verbose:$false -EA Ignore
    Import-Module $testModule -Force -Verbose:$false
}
BeforeAll {
    $testParams = @{
        ComputerName = 'GRPSDFRAN0049'
        Database     = 'PowerShell TEST'
        Table        = 'Tickets'
        TicketNr     = 1
        ScriptName   = 'Pester test'
    }

    $invokeTestParams = @{
        ServerInstance         = $testParams.ComputerName
        Database               = $testParams.Database
        TrustServerCertificate = $true
        QueryTimeout           = '1000'
        ConnectionTimeout      = '20'
        ErrorAction            = 'Stop'
        Verbose                = $false
    }
}
Describe 'Save-TicketInSqlHC' {
    $testCases = @(
        @{
            Name  = 'ServiceCountryCode'
            Value = 'FRA'
        }
        @{
            Name  = 'Requester'
            Value = 'TheRequester'
        }
        @{
            Name  = 'SubmittedBy'
            Value = 'TheSubmitter'
        }
        @{
            Name  = 'OwnedByTeam'
            Value = 'TEAM OPS'
        }
        @{
            Name  = 'OwnedBy'
            Value = 'TheOwner'
        }
        @{
            Name  = 'ShortDescription'
            Value = 'The short description'
        }
        @{
            Name  = 'Source'
            Value = 'TheSource'
        }
        @{
            Name  = 'Service'
            Value = 'TheService'
        }
        @{
            Name  = 'Category'
            Value = 'TheCategory'
        }
        @{
            Name  = 'SubCategory'
            Value = 'TheSubCategory'
        }
        @{
            Name  = 'IncidentType'
            Value = 'TheIncidentType'
        }
        @{
            Name  = 'Priority'
            Value = 3
        }
    )
    Context 'the mandatory parameters are' {
        It '<_>' -ForEach @(
            'KeyValuePair', 'TicketNr', 'PSCode', 'ScriptName'
        ) {
            (
                Get-Command -Name Save-TicketInSqlHC
            ).Parameters[$_].Attributes.Mandatory | Should -BeTrue
        }
    }
    It 'save ScriptName, TicketNr and PSCode' {
        $PSCode = New-PSCodeHC -CountryCode 'BNL'

        Save-TicketInSqlHC @testParams -PSCode $PSCode -KeyValuePair @{}

        $Actual = Invoke-Sqlcmd @invokeTestParams -Query "
                SELECT *
                FROM $($testParams.Table)
                WHERE PSCode = '$PSCode'"

        $Actual | Should -Not -BeNullOrEmpty
    }
    Context 'Save a new row with field' {
        It '<Name>' -ForEach $testCases {
            $PSCode = New-PSCodeHC -CountryCode 'BNL'

            Save-TicketInSqlHC @testParams -PSCode $PSCode -KeyValuePair @{$Name = $Value }

            $Actual = Invoke-Sqlcmd @invokeTestParams  -Query "
                SELECT *
                FROM $($testParams.Table)
                WHERE PSCode = '$PSCode'"

            $Actual.$Name | Should -BeExactly $Value
        }
    }
    Context 'Update an existing row with field' {
        BeforeAll {
            $PSCode = New-PSCodeHC -CountryCode 'BNL'
            Save-TicketInSqlHC @testParams -PSCode $PSCode -KeyValuePair @{}
        }
        It '<Name>' -ForEach $testCases {
            Save-TicketInSqlHC @testParams -PSCode $PSCode -KeyValuePair @{$Name = $Value } -Force

            $Actual = Invoke-Sqlcmd @invokeTestParams  -Query "
                SELECT *
                FROM $($testParams.Table)
                WHERE PSCode = '$PSCode'"

            $Actual.$Name | Should -BeExactly $Value
        }
    }
}
