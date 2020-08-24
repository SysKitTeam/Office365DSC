[CmdletBinding()]
param(
)
$M365DSCTestFolder = Join-Path -Path $PSScriptRoot `
                        -ChildPath "..\..\Unit" `
                        -Resolve
$CmdletModule = (Join-Path -Path $M365DSCTestFolder `
            -ChildPath "\Stubs\Microsoft365.psm1" `
            -Resolve)
$GenericStubPath = (Join-Path -Path $M365DSCTestFolder `
    -ChildPath "\Stubs\Generic.psm1" `
    -Resolve)
Import-Module -Name (Join-Path -Path $M365DSCTestFolder `
        -ChildPath "\UnitTestHelper.psm1" `
        -Resolve)

$Global:DscHelper = New-M365DscUnitTestHelper -StubModule $CmdletModule `
    -DscResource "EXOApplicationAccessPolicy" -GenericStubModule $GenericStubPath
Describe -Name $Global:DscHelper.DescribeHeader -Fixture {
    InModuleScope -ModuleName $Global:DscHelper.ModuleName -ScriptBlock {
        Invoke-Command -ScriptBlock $Global:DscHelper.InitializeScript -NoNewScope

        BeforeAll {
            $secpasswd = ConvertTo-SecureString "test@password1" -AsPlainText -Force
            $GlobalAdminAccount = New-Object System.Management.Automation.PSCredential ("tenantadmin", $secpasswd)

            Mock -CommandName Update-M365DSCExportAuthenticationResults -MockWith {
                return @{}
            }

            Mock -CommandName Get-M365DSCExportContentForResource -MockWith {

            }

            Mock -CommandName New-M365DSCConnection -MockWith {
                return "Credential"
            }

            Mock -CommandName Get-PSSession -MockWith {

            }

            Mock -CommandName Remove-PSSession -MockWith {

            }
        }

        # Test contexts
        Context -Name "Application Access Policy should exist. Application Access Policy is missing. Test should fail." -Fixture {
            BeforeAll {
                $testParams = @{
                    Identity           = "ApplicationAccessPolicy1"
                    AccessRight        = "DenyAccess"
                    AppID              = "3dbc2ae1-7198-45ed-9f9f-d86ba3ec35b5"
                    PolicyScopeGroupId = "Engineering Staff"
                    Description        = "Engineering Group Policy"
                    Ensure             = 'Present'
                    GlobalAdminAccount = $GlobalAdminAccount
                }

                Mock -CommandName Get-ApplicationAccessPolicy -MockWith {
                    return @{
                        Identity      = "DifferentApplicationAccessPolicy1"
                        AccessRight   = "DenyAccess"
                        AppID         = "3dbc2ae1-7198-45ed-9f9f-d86ba3ec35b5"
                        ScopeIdentity = "Engineering Staff"
                        Description   = "Engineering Group Policy"
                    }
                }

                Mock -CommandName Set-ApplicationAccessPolicy -MockWith {
                    return @{
                        Identity           = "ApplicationAccessPolicy1"
                        Description        = "Engineering Group Policy"
                        Ensure             = 'Present'
                        GlobalAdminAccount = $GlobalAdminAccount
                    }
                }
            }

            It 'Should return false from the Test method' {
                Test-TargetResource @testParams | Should -Be $false
            }

            It "Should call the Set method" {
                Set-TargetResource @testParams
            }

            It "Should return Absent from the Get method" {
                (Get-TargetResource @testParams).Ensure | Should -Be "Absent"
            }
        }

        Context -Name "Application Access Policy should exist. Application Access Policy exists. Test should pass." -Fixture {
            BeforeAll {
                $testParams = @{
                    Identity           = "ApplicationAccessPolicy1"
                    AccessRight        = "DenyAccess"
                    AppID              = "3dbc2ae1-7198-45ed-9f9f-d86ba3ec35b5"
                    PolicyScopeGroupId = "Engineering Staff"
                    Description        = "Engineering Group Policy"
                    Ensure             = 'Present'
                    GlobalAdminAccount = $GlobalAdminAccount
                }

                Mock -CommandName Get-ApplicationAccessPolicy -MockWith {
                    return @{
                        Identity      = "ApplicationAccessPolicy1"
                        AccessRight   = "DenyAccess"
                        AppID         = "3dbc2ae1-7198-45ed-9f9f-d86ba3ec35b5"
                        ScopeIdentity = "Engineering Staff"
                        Description   = "Engineering Group Policy"
                    }
                }
            }

            It 'Should return true from the Test method' {
                Test-TargetResource @testParams | Should -Be $true
            }

            It 'Should return Present from the Get Method' {
                (Get-TargetResource @testParams).Ensure | Should -Be "Present"
            }
        }

        Context -Name "Application Access Policy should exist. Application Access Policy exists, PolicyScopeGroupId mismatch. Test should fail." -Fixture {
            BeforeAll {
                $testParams = @{
                    Identity           = "ApplicationAccessPolicy1"
                    AccessRight        = "DenyAccess"
                    AppID              = "3dbc2ae1-7198-45ed-9f9f-d86ba3ec35b5"
                    PolicyScopeGroupId = "Engineering Staff"
                    Description        = "Engineering Group Policy"
                    Ensure             = 'Present'
                    GlobalAdminAccount = $GlobalAdminAccount
                }

                Mock -CommandName Get-ApplicationAccessPolicy -MockWith {
                    return @{
                        Identity      = "ApplicationAccessPolicy1"
                        AccessRight   = "DenyAccess"
                        AppID         = "3dbc2ae1-7198-45ed-9f9f-d86ba3ec35b5"
                        ScopeIdentity = "Finance Team"
                        Description   = "Engineering Group Policy"
                    }
                }

                Mock -CommandName Set-ApplicationAccessPolicy -MockWith {
                    return @{
                        Identity           = "ApplicationAccessPolicy1"
                        AccessRight        = "DenyAccess"
                        AppID              = "3dbc2ae1-7198-45ed-9f9f-d86ba3ec35b5"
                        PolicyScopeGroupId = "Engineering Staff"
                        Description        = "Engineering Group Policy"
                        Ensure             = 'Present'
                        GlobalAdminAccount = $GlobalAdminAccount
                    }
                }
            }

            It 'Should return false from the Test method' {
                Test-TargetResource @testParams | Should -Be $false
            }

            It "Should call the Set method" {
                Set-TargetResource @testParams
            }
        }

        Context -Name "ReverseDSC Tests" -Fixture {
            BeforeAll {
                $testParams = @{
                    GlobalAdminAccount = $GlobalAdminAccount
                }

                $ApplicationAccessPolicy = @{
                    Identity      = "ApplicationAccessPolicy1"
                    AccessRight   = "DenyAccess"
                    AppID         = "3dbc2ae1-7198-45ed-9f9f-d86ba3ec35b5"
                    ScopeIdentity = "Engineering Staff"
                    Description   = "Engineering Group Policy"
                }
                Mock -CommandName Get-ApplicationAccessPolicy -MockWith {
                    return $ApplicationAccessPolicy
                }
            }

            It "Should Reverse Engineer resource from the Export method when single" {
                Export-TargetResource @testParams
            }
        }
    }
}

Invoke-Command -ScriptBlock $Global:DscHelper.CleanupScript -NoNewScope