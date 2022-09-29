10.0.0.20210805
$version = "10.0.0.20210805"
<#
.SYNOPSIS 
    Chimera

.DESCRIPTION
    Fork of the Trident GUI 

.USAGE
    Right click on the program and select 'Run with Powershell'

.NOTES
    Version 10.0.0
    Author: James Brotosky
    Creation Date: 20210129
    Build Date: 20210805

.UPDATE NOTES
    Version 10.0.0 - Changed help button to open Chimera's usage documentation rather than a message box
                    Added WVU Medicine Icon to Chimera
                    Final release - Gave same level of access to everyone who runs the program 


    Version 9.0.0 - Auto open file location when new version of Chimera is active
                    Fixed Computer Management (RDP) bug

    Version 8.3.0 - Added PC Tech Support 

    Version 8.2.1 - Removed SSO as it didn't work consistantly 

    Version 8.2.0 - Added SSO to Remote Desktop User Modification

    Version 8.1.0 - Change search element to only search WVUHS domain

    Version 8.0.0 - Visual Element Overhaul

    Version 7.1.0 - Added Email box to Associate Help Desk level access 

    Version 7.0.0 - Added System Analyst access

    Version 6.1.1 - Bug Fix: Field Highlighter highlighted incorrect fields 

    Version 6.1.0 - Removed 'Insert Username Here' Text
                    Made keyboard cursor focus on username input box upon launch

    Version 6.0.0 - Major UI Changes
                    Remove PC Tech Functions
                    Code Overhaul for efficiency
                    Sanitized user input
                    Removed COVID Portal Buttons
                    Color Coated Account Status and Password Status Text Boxes

    Version 5.5.0 - Changed update managing logic to look in **REDACTED** instead of my personal folder. Some helpdesk staff don't have access to my folder
                    Added connection testing for computer file explorer, so it doesn't take so long to error out if it can't reach the PC

    Version 5.4.0 - Added SEC-COVID-19 Check
                    Faster Loading

    Version 5.3.1 - Enabled Visual Styles

    Version 5.3.0 - Added new Office 365 License field
                    Gave T1 Techs Windows/IGEL Check field 

    Version 5.2.0 - Updated to include F3 licenses for transition from F1 licenses

    Version 5.1.0 - Bug fix for rebooting computers
                    Removed some errors on startup

    Version 5.0.0 - Faster Loading
                    Check for updates instead of tamper evidence
                    Gave tier 1 techs reboot with auto remote in
                    Cleaned up distribution group and deletegate boxes
                    Removed firstname and lastname boxes since display name is available
                    Changed PC Tech GUI and added room for more features in the future

    Version 4.1.0 - Added Delegate of mailboxes and manager of distribution groups information to senior tech function

    Version 4.0.0 - Added PC Tech functions

    Version 3.0.0 - Added Remote Desktop User Modification to T3 level access

    Version 2.1.0 - Added WVUM-O365F1Teams to Client check

    Version 2.0.0 - Upon request, gave more access to Tier 2 HD techs 

    Version 1.0.0 - First Release

    Version 0.0.1 - Initial Build

.REMINDER
    When pushing new version, don't forget to uncomment the update check portion.
#>
$devFlag = 0
if ( $devFlag -eq "0" ) {
    $ErrorActionPreference = "SilentlyContinue"
	$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
	add-type -name win -member $t -namespace native
	[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 0)

    #Code Integrity Check
    $versionNumber = Get-Content '**REDACTED**' | Select -Index 0
    if ( $versionNumber -ne $version ) { 
        ii "**REDACTED**"
        Start-Sleep 1
        [Windows.Forms.MessageBox]::Show("Newer version of Chimera found:`n`nPlease copy and paste Chimera from the open window to your preferred run location.", "Update", [Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Information)
        Write-Host $versionNumber
        $objLOADINGFORM.Close()
        exit
    } else {
    }
    $currentFile = $MyInvocation.MyCommand.Source
    $SourceFile = Get-FileHash $currentFile -Algorithm SHA1
    $DestFile = Get-FileHash '**REDACTED**' -Algorithm SHA1 
    if ( $SourceFile.Hash -ne $DestFile.Hash ) {
        [Windows.Forms.MessageBox]::Show("This is not a geunine copy of Chimera.`nThis is being reported to the program developer.", "Exiting Application", [Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Warning)
        $objLOADINGFORM.Close()
        exit
    } else {
    }
} else {
    $ErrorActionPreference = "Continue"
	$t = '[DllImport("user32.dll")] public static extern bool ShowWindow(int handle, int state);'
	add-type -name win -member $t -namespace native
	[native.win]::ShowWindow(([System.Diagnostics.Process]::GetCurrentProcess() | Get-Process).MainWindowHandle, 3)
}
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.VisualSytles")
[Windows.Forms.Application]::EnableVisualStyles()
#Loading Form
$objLoadingForm = New-Object System.Windows.Forms.Form
$objLoadingForm.StartPosition = "CenterScreen"
$objLoadingForm.Text = "Loading Chimera"
$objLoadingForm.Size = New-Object System.Drawing.Size(500,250)
$objLoadingFormLabel = New-Object System.Windows.Forms.Label
$objLoadingFormLabel.Text = "Importing Active Directory..."
$objLoadingFormLabel.Location = New-Object System.Drawing.Size(175,50)
$objLoadingFormLabel.Size = New-Object System.Drawing.Size(250,23)
$objLoadingFormProgressAmount = 33
$objLoadingFormProgressBar = New-Object System.Windows.Forms.ProgressBar
$objLoadingFormProgressBar.Location = New-Object System.Drawing.Size(10,150)
$objLoadingFormProgressBar.Size = New-Object System.Drawing.Size(460,23)
$objLoadingFormProgressBar.Value = $objLoadingFormProgressAmount
$objLoadingFormProgressBar.Style = "Continuous"
$objChimeraFormIconPath = '**REDACTED**'
$objLoadingForm.Icon = New-Object System.Drawing.Icon($objChimeraFormIconPath)
$objLoadingForm.Controls.Add($objLoadingFormLabel)
$objLoadingForm.Controls.Add($objLoadingFormProgressBar)
$objLoadingForm.Show()
$objLoadingForm.Refresh()
Start-Sleep -Seconds 1
#Access Level Logic
Import-Module ActiveDirectory
$objLoadingFormLabel.Text = "Checking for updates..."
$objLoadingFormProgressAmount = 66
$objLoadingFormProgressBar.Value = $objLoadingFormProgressAmount
$objLoadingForm.Refresh()
Start-Sleep -Milliseconds 700
$objLoadingFormLabel.Text = "Checking Access Level..."
$optionalT3 = Get-Content "**REDACTED**"
$optionalT2 = Get-Content "**REDACTED**"
$currentUser = Get-ADUser $env:Username -Properties *
$objLoadingFormProgressAmount = 99
$objLoadingFormProgressBar.Value = $objLoadingFormProgressAmount
$objLoadingForm.Refresh()
Start-Sleep -Milliseconds 200
$objLoadingFormLabel.Text = "Finished, Launching Chimera"
$objLoadingFormProgressAmount = 100
$objLoadingFormProgressBar.Value = $objLoadingformProgressAmount
$objLoadingForm.Refresh()
Start-Sleep -Milliseconds 200
$objLoadingForm.Close()
Function FindADUser ( $FindADUserName ) {
if ( $FindADUserName -ne "" ) {
    $objCreationDateTextBox.Clear()
    $objAccountStatusTextBox.Clear()
    $objMyIdentityTextBox.Clear()
    $objDisplayNameTextBox.Clear()
    $objDistinguishedNameTextBox.Clear()
    $objEmailTextBox.Clear()
    $objEnterPriseIDTextBox.Clear()
    $objJobTitleTextBox.Clear()
    $objLastBadPasswordTextBox.Clear()
    $objLastLogonTextBox.Clear()
    $objOrganizationalUnitTextBox.Clear()
    $objOffice365TextBox.Clear()
    $objPasswordLastSetTextBox.Clear()
    $objPasswordStatusTextBox.Clear()
    $objPersonalDriveTextBox.Clear()
    $objVPNUserTextBox.Clear()
    $objMobileEmailTextBox.Clear()
    $objWebClientTextBox.Clear()
    $objMemberOfTextBox.Items.Clear()
    $objSelectedUserLabel.BackColor = ""
    $objAccountStatusTextBox.BackColor = ""
    $objPasswordStatusTextBox.BackColor = ""
    $userSearch = $objTextBox.Text
    $userSearchSani = $usersearch.Trim()
    $objChimeraForm.Refresh()
    $objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $domainList = @($objForest.Domains | Select-Object Name)
    $domains = $domainList | ForEach {$_.Name}
    Try { 
        (Get-ADUser -Identity $userSearchSani -Server **REDACTED**)
        $objSelectedUserLabel.Text = "Searching the **REDACTED** Domain"
        $userFound = "True"
        $Domain = "**REDACTED**"
        $objChimeraForm.Refresh()
    } Catch {
        $userFound = "False"
    }
    }
    if ( $userFound -eq "True" ) {
        $user = Get-ADUser -identity $userSearchSani -Server $Domain -Properties *
        $objSelectedUserLabel.Text = $user.DisplayName
        #$objSelectedUserLabel.BackColor = "Red"
        #$objSelectedUserLabel.ForeColor = "Black"
        $objCreationDateTextBox.Text = $user.Created
        $objAccountStatusTextBox.Text = $user.Enabled
        $objMyIdentityTextBox.Text = $user.ExtensionAttribute3
        $objDisplayNameTextBox.Text = $user.DisplayName
        $objDistinguishedNameTextBox.Text = $user.DistinguishedName
        $objEmailTextBox.Text = $user.Mail
        $objEnterPriseIDTextBox.Text = $user.extensionAttribute13
        $objJobTitleTextBox.Text = $user.Title
        $objLastBadPasswordTextBox.Text = $user.LastBadPasswordAttempt
        $objLastLogonTextBox.text = $user.LastLogonDate
        $objOrganizationalUnitTextBox.Text = $user.CanonicalName
        if ( $user.TargetAddress -eq $null ) {
            $objOffice365TextBox.Text = "No"
        } else {
            $objOffice365TextBox.Text = "Yes"
        }
        $objPasswordLastSetTextBox.Text = $user.PasswordLastSet
        $objPasswordStatusTextBox.Text = $user.PasswordExpired
        $objPersonalDriveTextBox.Text = $user.HomeDirectory
        $objMemberOfTextBoxList = $user.memberof -replace '^CN=([^,]+).+$','$1' | Sort-Object
        $objMemberOfTextBox.Items.AddRange($objMemberOfTextBoxList)
        if ( $objMemberOfTextBoxList -Match "**REDACTED**" ) {
            $objVPNUserTextBox.Text = "Yes"
        } else {
            $objVPNUserTextBox.Text = "No"
        }
        if ( $objMemberOfTextBoxList -Match "wvum-mam" ) {
            $objMobileEmailTextBox.Text = "Yes"
        } else {
            $objMobileEmailTextBox.Text = "No"
        }
        if ($objMemberOfTextBoxList -Match "WVUM-O365F1Office" ) { 
            $objWebClientTextBox.Text = "Web Client Only"
        } elseif ( $objMemberOfTextBoxList -Match "WVUM-O365E3Office" ) { 
            $objWebClientTextBox.Text = "Web Client and Full Client"
        } elseif ($objMemberOfTextBoxList -Match "WVUM-O365E3Teams" ) {
            $objWebClientTextBox.Text = "Web Client and Full Client"
        } elseif ($objMemberOfTextBoxList -Match "WVUM-O365F1Teams" ) {
            $objWebClientTextBox.Text = "Web Client Only"
        } elseif ( $objMemberOfTextBoxList -Match "WVUM-O365F3Office" ) {
            $objWebClientTextBox.Text = "Web Client Only"
        } elseif ( $objMemberOfTextBoxList -Match "WVUM-O365F3Teams" ) { 
            $objWebClientTextBox.Text = "Web Client Only"
        } else {
            $objWebClientTextBox.Text = "Undefined or Abnormal License"
        }
        <#
            $delegates = $user.publicDelegatesBL
            $delegatesRegEx = $delegates -replace '^CN=([^,]+).+$','$1' | Sort-Object
            $objDelegatesListListBox.Items.AddRange($delegatesRegEx)
            $distributionGroups = $user.managedObjects
            $distributionGroupsRegEx = $distributionGroups -replace '^CN=([^,]+).+$','$1' | Sort-Object
            $objDistributionGroupListBox.Items.AddRange($distributionGroupsRegEx)
        #>
        if ( $objAccountStatusTextBox.Text -eq "True" ) {
            $objAccountStatusTextBox.BackColor = "LightGreen"
        } elseif ( $objAccountStatusTextBox.Text -eq "False" ) { 
            $objAccountStatusTextBox.BackColor = "Red"
        } else { 
            $objAccountStatusTextBox.BackColor = "White"
        }
        if ( $objPasswordStatusTextBox.Text -eq "True" ) { 
            $objPasswordStatusTextBox.BackColor = "Red"
        } elseif ( $objPasswordStatusTextBox.Text -eq "False" ) { 
            $objPasswordStatusTextBox.BackColor = "LightGreen"
        } else {
            $objPasswordStatusTextBox.BackColor = "White"
        }
    } else {
        $objSelectedUserLabel.Text = "User not found..."
        $objSelectedUserLabel.BackColor = ""
        $objAccountStatusTextBox.BackColor = ""
        $objPasswordStatusTextBox.BackColor = ""
        $objOffice365TextBox.Text = ""
        $objMobileEmailTextBox.Text = ""
        $objVPNUserTextBox.Text = ""
        $objWebClientTextBox.Text = ""
        $objChimeraForm.Refresh()
    }
}           



<#
Function AttributeInfo ( $user ) {
    if ( $user -ne $null ) {
        $objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
        $DomainList = @($objForest.Domains | Select-Object Name)
        Write-Host $DomainList
        $Domains = $DomainList | ForEach {$_.Name}
        ForEach ( $Domain in ( $Domains ) ) {
            if ( -Not ( Get-ADGroup -Identity $objMemberofListBox.SelectedItem -Server $Domain ) ) {
            } else {
                $attinfo = Get-ADGroup -Identity $objMemberofListBox.SelectedItem -Properties * -Server $Domain
                $attInfoName = $attinfo.Name
                $attInfoDescription = $attinfo.Description
                [Windows.Forms.MessageBox]::Show("$attInfoDescription", "$attInfoName", [Windows.Forms.MessageBoxButtons]::OK,[Windows.Forms.MessageBoxIcon]::Information)
            }
        }
    }
}
#>
Function ComputerFileExplorer ( $pc ) {
    $pcTag = $objPCTagTextBox.Text
    $pcTagSani = $pcTag.Trim()
    if (( $pcTagSani -eq "" ) -or ( $pcTagSani -eq $null)) {
        [Windows.Forms.MessageBox]::Show("Computer Tag Cannot Be Blank", "Empty Computer Tag", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        $pc = $pcTagSani
        ii "\\$pc\C$"
    }
}

Function LAPS ( $password ) {
    $pcTag = $objPCTagTextBox.Text
    $pcTagSani = $pcTag.Trim()
    if (( $pcTagSani -eq "" ) -or ( $pcTagSani -eq $null )) {
        [Windows.Forms.MessageBox]::Show("Computer Tag Cannot Be Blank", "Empty Computer Tag", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning)
    } else {
    $objForest = [System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest()
    $objDomainList = @($objForest.Domains | Select-Object Name)
    $objDomains = $objDomainList | ForEach {$_.Name}
        ForEach ( $domain in $objDomains ) {
            if ( Get-ADComputer $pcTagSani -Server $domain ) {
                $objPCTagString = $pcTagSani | Out-String
                $password = Get-ADComputer -Identity $pcTagSani -Server $domain -Properties * | Select -ExpandProperty ms-Mcs-AdmPwd
                $objLapsForm = New-Object System.Windows.Forms.Form
                $objLapsForm.Text = "Laps Password"
                $objLapsForm.Size = New-Object System.Drawing.Size(300,120)
                $objLapsForm.StartPosition = "CenterScreen"
                $objLapsForm.KeyPreview = $True
                $objLapsFormLabel = New-Object System.Windows.Forms.Label
                $objLapsFormLabel.Text = "Laps Password for: $objPCTagString"
                $objLapsFormLabel.Location = New-Object System.Drawing.Size(10,10)
                $objLapsFormLabel.AutoSize = $True
                $objLapsForm.Controls.Add($objLapsFormLabel)
                $objLapsFormText = New-Object System.Windows.Forms.TextBox
                $objLapsFormText.Text = $password 
                $objLapsFormText.Location = New-Object System.Drawing.Size(20,50)
                $objLapsFormText.Size = New-Object System.Drawing.Size(250,10)
                $objLapsFormText.Focus()
                $objLapsForm.Controls.Add($objLapsFormText)
                [void] $objLapsForm.Show()
                $objLapsForm.Refresh()
            }
        }
    }
}

Function REMOTE ( $pc ) {
    $pcTag = $objPCTagTextBox.Text
    $pcTagSani = $pcTag.Trim()
    if (( $pcTagSani -eq "" ) -or ( $pcTagSani -eq $null )) {
        [Windows.Forms.MessageBox]::Show("Computer Tag Cannot Be Blank", "Empty Computer Tag", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning)
    } else {
    $testConnection = Test-Connection -Count 1 $pcTagSani
        if ( $testConnection -eq "False" ) {
            [Windows.Forms.MessageBox]::Show("Computer is unresponsive, not connected to the network, or is an IGEL Machine", "Connection Failed", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Stop)
        } else {
            $path1 = Test-Path "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386"
            $path2 = Test-Path "C:\Program Files (x86)\Microsoft Configuration Manager\bin\i386"
            $path3 = Test-Path "C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\i386"
            $path4 = Test-Path "C:\Program Files\Microsoft Configuration Manager\bin\i386"
            if ( $path1 -Match "True" ) {
                & 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe' $pcTagSani
                Start-Sleep -s 3
                clear
            } elseif ( $path2 -Match "True" ) {
                & 'C:\Program Files (x86)\Microsoft Configuration Manager\bin\i386\CmRcViewer.exe' $pcTagSani
                Start-Sleep -s 3
                clear
            } elseif ( $path3 -Match "True" ) {
                & 'C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe' $pcTagSani
                Start-Sleep -s 3
                clear
            } elseif ( $path4 -Match "True" ) {
                & 'C:\Program Files\Microsoft Configuration Manager\bin\i386\CmRcViewer.exe' $pcTagSani
                Start-Sleep -s 3
                clear
            }
        }
    }
}

Function REBOOTREMOTE ( $pc ) {
    $pcTag = $objPCTagTextBox.Text
    $pcTagSani = $pcTag.Trim()
    if (( $pcTagSani -eq "" ) -or ( $pcTagSani -eq $null )) {
        [Windows.Forms.MessageBox]::Show("Computer Tag Cannot Be Blank", "Empty Computer Tag", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        if (( $pcTagSani -eq "" ) -or ( $pcTagSani -eq $null )) {
            [Windows.Forms.MessageBox]::Show("Computer Tag Cannot Be Blank", "Error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error)
        } else {
            $pc = $pcTagSani
            $confirm = [Windows.Forms.MessageBox]::Show("Are you sure you want to reboot $pc`?","Confirm", [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Question)
            if ( $confirm -eq "Yes" ) {
                $pcFull = "\\" + $pc
                $shutDownMessage = [Windows.Forms.MessageBox]::Show("Shutting Down $pc", "Shutting Down", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
                Restart-Computer -ComputerName $pc -Force
                Start-Sleep -Seconds 5
                Do {
                    $shutDown = Test-Connection $pc -Count 1 -Quiet
                    Start-Sleep -Seconds 2
                } Until ( $shutDown -eq 'False' )
                    $restartMessage = [Windows.Forms.MessageBox]::Show("Waiting for computer to restart, this might take awhile", "Restarting Computer", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
                    Start-Sleep -Seconds 5
                Do {
                    $restart = Test-Connection $pc -Count 1 -Quiet
                    Start-Sleep -Seconds 2
                } Until ( $restart -eq "True" ) 
                $remoteMessage = [Windows.Forms.MessageBox]::Show("Host is now live, would you like to remote in?", "Host Live", [Windows.Forms.MessageBoxButtons]::YesNo, [Windows.Forms.MessageBoxIcon]::Information)
                if ( $remoteMessage -eq "Yes" ) {
                    $path1 = Test-Path "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386"
                    $path2 = Test-Path "C:\Program Files (x86)\Microsoft Configuration Manager\bin\i386"
                    $path3 = Test-Path "C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\i386"
                    $path4 = Test-Path "C:\Program Files\Microsoft Configuration Manager\bin\i386"
                    if ( $path1 -Match "True" ) {
                        & 'C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe' $objpctagTextBox.Text
                        Start-Sleep -s 3
                        clear
                    } elseif ( $path2 -Match "True" ) {
                        & 'C:\Program Files (x86)\Microsoft Configuration Manager\bin\i386\CmRcViewer.exe' $objpctagTextBox.Text
                        Start-Sleep -s 3
                        clear
                    } elseif ( $path3 -Match "True" ) {
                        & 'C:\Program Files\Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe' $objpctagTextBox.Text
                        Start-Sleep -s 3
                        clear
                    } elseif ( $path4 -Match "True" ) {
                        & 'C:\Program Files\Microsoft Configuration Manager\bin\i386\CmRcViewer.exe' $objpctagTextBox.Text
                        Start-Sleep -s 3
                        clear
                    }
                }
            }
        }
    }
}

Function WINDOWSIGELCHECK ( $pc ) {
    $pcTag = $objPCTagTextBox.Text
    $pcTagSani = $pcTag.Trim() 
    if (( $pcTagSani -eq "" ) -or ( $pcTagSani -eq "null" )) {
        [Windows.Forms.MessageBox]::Show("Computer Tag Cannot Be Blank", "Error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        $a = 0
        $b = 0
        $pc = $pcTagSani
        Write-Host $pcfull
        $check1 = Test-Path \\$pc\c$
        if ( $check1 -eq 'True' ) { 
            $a = $a + 33
        } else { 
            $b = $b + 33
        }
        Try { 
            Get-ADComputer $pc
            $a = $a + 33
        } Catch { 
            $ErrorMessage = $_.Exception.Message
            $b = $b + 33
        }
        if ( $a -eq 99 ) { 
            $a = $a + 1
        } elseif ( $b -eq 99 ) { 
            $b = $b + 1
        } else { 
        }
        $objWINDOWSIGELCHECKForm = New-Object System.Windows.Forms.Form
        $objWINDOWSIGELCHECKForm.Text = "Operating System Type Confidence"
        $objWINDOWSIGELCHECKForm.Size = New-Object System.Drawing.Size(500,250)
        $objWINDOWSIGELCHECKForm.StartPosition = "CenterScreen"
        $objWINDOWSIGELCHECKForm.Add_KeyDown({if ( $_.KeyDown -eq "Escape" ) { $objWINDOWSIGELCHECKForm.Close()}})
        $objWINDOWSIGELCHECKFormLabel = New-Object System.Windows.Forms.Label 
        $objWINDOWSIGELCHECKFormLabel.Text = " Operating System Confidence"
        $objWINDOWSIGELCHECKFormLabel.Location = New-object System.Drawing.Size(125,5)
        $objWINDOWSIGELCHECKFormLabel.Size = New-Object System.Drawing.Size(300,30)
        $objWINDOWSIGELCHECKFormLabel.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
        $objWINDOWSIGELCHECKFormLabelComputerName = New-Object System.Windows.Forms.Label
        $objWINDOWSIGELCHECKFormLabelComputerName.Text = "Computer Name: $pc"
        $objWINDOWSIGELCHECKFormLabelComputerName.Location = New-Object System.Drawing.Size(175,37)
        $objWINDOWSIGELCHECKFormlabelComputerName.Font = New-Object System.Drawing.Font("Segoe UI",7,[System.Drawing.FontStyle]::Regular)
        $objWINDOWSIGELCHECKFormlabelComputerName.Size = New-Object System.Drawing.Size(200,15)
        $objWINDOWSIGELCHECKFormLabelWindows = New-Object System.Windows.Forms.Label
        $objWINDOWSIGELCHECKFormLabelWindows.Text = "Windows:"
        $objWINDOWSIGELCHECKFormLabelWindows.Location = New-Object System.Drawing.Size(10,60)
        $objWINDOWSIGELCHECKFormLabelWindows.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
        $objWINDOWSIGELCHECKFormLabelWindowsPercent = New-Object System.Windows.Forms.Label
        $objWINDOWSIGELCHECKFormLabelWindowsPercent.Text = "$a%" 
        $objWINDOWSIGELCHECKFormLabelWindowsPercent.Location = New-Object System.Drawing.Size(445,61)
        $objWINDOWSIGELCHECKFormLabelWindowsPercent.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
        $objWINDOWSIGELCHECKFormLabelWindowsProgressBar = New-Object System.Windows.Forms.ProgressBar
        $objWINDOWSIGELCHECKFormLabelWindowsProgressBar.Name = "Test"
        $objWINDOWSIGELCHECKFormLabelWindowsProgressBar.Value = $a
        $objWINDOWSIGELCHECKFormLabelWindowsProgressBar.Style = "Continuous"
        $objWINDOWSIGELCHECKFormLabelWindowsProgressBar.Location = New-Object System.Drawing.Size(110,63)
        $objWINDOWSIGELCHECKFormLabelWindowsProgressBar.Size = New-Object System.Drawing.Size(300,20)
        $objWINDOWSIGELCHECKFormLabelLinux = New-Object System.Windows.Forms.Label
        $objWINDOWSIGELCHECKFormLabelLinux.Text = "IGEL:"
        $objWINDOWSIGELCHECKFormLabelLinux.Location = New-Object System.Drawing.Size(43,100)
        $objWINDOWSIGELCHECKFormLabelLinux.Size = New-Object System.Drawing.Size(50,30)
        $objWINDOWSIGELCHECKFormLabelLinux.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
        $objWINDOWSIGELCHECKFormLabelLinuxPercent = New-Object System.Windows.Forms.Label
        $objWINDOWSIGELCHECKFormLabelLinuxPercent.Text = "$b%"
        $objWINDOWSIGELCHECKFormLabelLinuxPercent.Location = New-Object System.Drawing.Size(445,102)
        $objWINDOWSIGELCHECKFormLabelLinuxPercent.Font = New-Object System.Drawing.Font("Segoe UI",12,[System.Drawing.FontStyle]::Regular)
        $objWINDOWSIGELCHECKFormLabelLinuxProgressBar = New-Object System.Windows.Forms.ProgressBar
        $objWINDOWSIGELCHECKFormLabelLinuxProgressBar.Value = $b
        $objWINDOWSIGELCHECKFormLabelLinuxProgressBar.Style = "Continuous"
        $objWINDOWSIGELCHECKFormLabelLinuxProgressBar.Location = New-Object System.Drawing.Size(110,103)
        $objWINDOWSIGELCHECKFormLabelLinuxProgressBar.Size = New-Object System.Drawing.Size(300,20)
        $objWINDOWSIGELCHECKFormLabelWarning = New-Object System.Windows.Forms.Label
        $objWINDOWSIGELCHECKFormlabelWarning.Text = "WARNING: This feature is still experimental and in development and only checks against two factors, making it not 100% reliable. Currently it checks against to see if C$ exists, and if the computer exists in Active Directory. Please be advised that a computer that was recently made into an IGEL machine may still show up in AD and give false confidence to the test. If you have any ideas of how to check for the operating system type further, please reach out to James Brotosky."
        $objWINDOWSIGELCHECKFormLabelWarning.Location = New-Object System.Drawing.Size(10,135)
        $objWINDOWSIGELCHECKFormLabelWarning.Size = New-Object System.Drawing.Size(475,65)
        $objWINDOWSIGELCHECKFormLabelWarning.Font = New-Object System.Drawing.Font("Segoe UI",7,[System.Drawing.FontStyle]::Regular)
        $objWINDOWSIGELCHECKForm.Controls.Add($objWINDOWSIGELCHECKFormLabel)
        $objWINDOWSIGELCHECKForm.Controls.add($objWINDOWSIGELCHECKFormLabelComputerName)
        $objWINDOWSIGELCHECKForm.Controls.Add($objWINDOWSIGELCHECKFormLabelWindows)
        $objWINDOWSIGELCHECKForm.Controls.Add($objWINDOWSIGELCHECKFormLabelWindowsPercent)
        $objWINDOWSIGELCHECKForm.Controls.Add($objWINDOWSIGELCHECKFormLabelWindowsProgressBar)
        $objWINDOWSIGELCHECKForm.Controls.Add($objWINDOWSIGELCHECKFormLabelLinux)
        $objWINDOWSIGELCHECKForm.Controls.Add($objWINDOWSIGELCHECKFormLabelLinuxPercent)
        $objWINDOWSIGELCHECKForm.Controls.Add($objWINDOWSIGELCHECKFormLabelLInuxProgressBar)
        $objWINDOWSIGELCHECKForm.Controls.Add($objWINDOWSIGELCHECKFormLabelWarning)
        $objWINDOWSIGELCHECKForm.Add_Shown({$objWINDOWSIGELCHECKForm.Activate(); $objWINDOWSIGELCHECKForm.Focus()})
        [void]$objWINDOWSIGELCHECKForm.ShowDialog()
    }
} 

Function REMOTEDESKTOPUSER ( $pc ) {
    $pcTag = $objPCTagTextBox.Text
    $pcTagSani = $pcTag.Trim()
    if (( $pcTagSani -eq "" ) -or ( $pcTagSani -eq $null )) {
        [Windows.Forms.MessageBox]::Show("Computer Tag Cannot Be Blank", "Error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Warning)
    } else {
        $pc = $pcTagSani
        $objREMOTEDESKTOPUSERForm = New-Object System.Windows.Forms.Form
        $objREMOTEDESKTOPUSERForm.Text = "Remote Desktop User Modification"
        $objREMOTEDESKTOPUSERForm.Size = New-Object System.Drawing.Size(500,250)
        $objREMOTEDESKTOPUSERForm.StartPosition = "CenterScreen"
        $objREMOTEDESKTOPUSERFormPCTaginfo = New-Object System.Windows.Forms.Label
        $objREMOTEDESKTOPUSERFormPCTaginfo.Text = "Remote Desktop User Modification for $pc"
        $objREMOTEDESKTOPUSERFormPCTaginfo.Location = New-Object System.Drawing.Size(125,25)
        $objREMOTEDESKTOPUSERFormPCTaginfo.Size = New-Object System.Drawing.Size(250,23)
        $objREMOTEDESKTOPUSERUsername = New-Object System.Windows.Forms.TextBox
        $objREMOTEDESKTOPUSERUsername.Location = New-Object System.Drawing.Size(110,55)
        $objREMOTEDESKTOPUSERUsername.Size = New-Object System.Drawing.Size(260,23)
        $objREMOTEDESKTOPUSERUsername.Focus()
        $objREMOTEDESKTOPUSERUsername.Text = "Input Username"
        $objREMOTEDESKTOPUSERButton = New-Object System.Windows.Forms.Button
        $objREMOTEDESKTOPUSERButton.Location = New-Object System.Drawing.Size(175,85)
        $objREMOTEDESKTOPUSERButton.Size = New-Object System.Drawing.Size(120,75)
        $objREMOTEDESKTOPUSERButton.Add_Click({(REMOTEDESKTOPUSERMODIFICATION)})
        $objREMOTEDESKTOPUSERButton.Text = "Add As Remote Desktop User"
        $objREMOTEDESKTOPUSERForm.Controls.Add($objREMOTEDESKTOPUSERButton)
        $objREMOTEDESKTOPUSERForm.Controls.Add($objREMOTEDESKTOPUSERFormPCTaginfo)
        $objREMOTEDESKTOPUSERForm.Controls.Add($objREMOTEDESKTOPUSERUsername)
        [void]$objREMOTEDESKTOPUSERForm.ShowDialog()
    }
} 

Function REMOTEDESKTOPUSERMODIFICATION ( $pc ) {
    $credential = Get-Credential
    $pcTag = $objPCTagTextBox.Text
    $pcTagSani = $pcTag.Trim()
    $a = 0
    $pc = $pcTagSani
    $User = $objREMOTEDESKTOPUSERUsername.Text
    $objREMOTEDESKTOPUSERModificationForm = New-Object System.Windows.Forms.Form
    $objREMOTEDESKTOPUserModificationForm.Text = "Status"
    $objREMOTEDESKTOPUSERModificationForm.StartPosition = "CenterScreen"
    $objREMOTEDESKTOPUSERModificationForm.Size = New-Object System.Drawing.Size(500,175)
    $objREMOTEDESKTOPUSERModificationFormLabel = New-Object System.Windows.Forms.Label
    $a = 33
    $objREMOTEDESKTOPUSERModificationFormLabel.Text = "Connecting to $pc..."
    $objREMOTEDESKTOPUSERModificationFormLabel.Location = New-Object System.Drawing.Size(175,35)
    $objREMOTEDESKTOPUSERModificationFormLabel.Size = New-Object System.Drawing.Size(250,23)
    $objREMOTEDESKTOPUSERModificationFormProgressBar = New-Object System.Windows.Forms.ProgressBar
    $objREMOTEDESKTOPUSERModificationFormProgressBar.Value = $a
    $objREMOTEDESKTOPUSERModificationFormProgressBar.Style = "Continuous"
    $objREMOTEDESKTOPUSERModificationFormProgressBar.Location = New-Object System.Drawing.Size(10,75)
    $objREMOTEDESKTOPUSERModificationFormProgressBar.Size = New-Object System.Drawing.Size(460,23)
    $objREMOTEDESKTOPUSERModificationForm.Controls.Add($objREMOTEDESKTOPUSERModificationFormLabel)
    $objREMOTEDESKTOPUSERModificationForm.Controls.Add($objREMOTEDESKTOPUSERModificationFormProgressBar)
    $objREMOTEDESKTOPUSERModificationForm.Show() | Out-Null
    $objREMOTEDESKTOPUSERModificationForm.Refresh()
    $Group = "Remote Desktop Users"
    $GroupObj = [ADSI]"WinNT://$pc/$Group,group"
    $MembersObj = @($GroupObj.psbase.Invoke("Members"))
    $a = 66
    $objREMOTEDESKTOPUSERModificationFormLabel.Text = "Attempting to add $user..."
    $objREMOTEDESKTOPUSERModificationFormProgressBar.Value = $a
    $objREMOTEDESKTOPUSERModificationForm.Refresh()
    $Members = ($MembersObj | ForEach {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)})
    if ( $Members -Contains $User ) { 
        [Windows.Forms.MessageBox]::Show("$User exists in $group, exiting...", "Remote Desktop Users and Groups", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
        $objREMOTEDESKTOPUSERForm.Close()
        $objREMOTEDESKTOPUSERModificationForm.Close()
        exit
    } else {
        Invoke-Command -ComputerName $pc -ScriptBlock { Add-LocalGroupMember -Group "Remote Desktop Users" -Member $Using:User } -Credential $credential
    }
    Remove-Module ActiveDirectory
    Import-Module ActiveDirectory
    $a = 99
    $objREMOTEDESKTOPUSERModificationFormLabel.Text = "Checking if adding user was successful..."
    $objREMOTEDESKTOPUSERModificationFormProgressBar.Value = $a
    $objREMOTEDESKTOPUSERModificationForm.Refresh()
    $Group = "Remote Desktop Users"
    $GroupObj = [ADSI]"WinNT://$pc/$group,group"
    $MembersObj = @($GroupObj.psbase.Invoke("Members"))
    $Members = ($MembersObj | ForEach {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)})
    sleep 5
    if ( $Members -Contains $User ) { 
        [Windows.Forms.MessageBox]::Show("$user was successfully added to Remote Desktop Users", "Success", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
        $objREMOTEDESKTOPUSERModificationForm.Close()
        $objREMOTEDESKTOPUSERForm.Close()
    } else {
        [Windows.Forms.MessageBox]::Show("Modification was Unsuccessfull, Please attempt to add the user manually", "Error", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error)
        $objREMOTEDESKTOPUSERModificationForm.Close()
        $objREMOTEDESKTOPUSERForm.Close()
    }
}

Function HELP () {
    ii "**REDACTED**"
    #[Windows.Forms.MessageBox]::Show("Welcome to Chimera!`n`nTo use:`n`nType in the username you wish to look up information for and select the search button.`n`nFor Computer tag functions:`n`nPlease type in a computer tag in the corresponding box and click an option to enact on the computer.", "Chimera Usage", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
}

Function WHATSNEW () {
    [Windows.Forms.MessageBox]::Show("Chimera - $Version`n`nWhats new in this update:`n`n- Auto open Chimera file location when new update is found`n`n- RDP Management Bug Fix", "Chimera $Version", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Information)
}
    $objChimeraForm = New-Object System.Windows.Forms.Form
    $objChimeraForm.Text = "Chimera - Version: $version"
    $objChimeraForm.AutoSize = $True
    $objChimeraForm.StartPosition = "CenterScreen"
    $objChimeraForm.KeyPreview = $True
    $objChimeraFormIconPath = '**REDACTED**'
    $objChimeraForm.Icon = New-Object System.Drawing.Icon($objChimeraFormIconPath)
    $objFindButton = New-Object System.Windows.Forms.Button
    $objFindButton.Location = New-Object System.Drawing.Size(10,55)
    $objFindButton.Size = New-Object System.Drawing.Size(73,23)
    $objFindButton.Text = "Search"
    $objFindButton.Add_Click({FindADUser $objTextBox.Text})
    $objInfoField = New-Object System.Windows.Forms.Label
    $objInfoField.AutoSize = $True
    $objInfoField.Name = "Usage"
    $objInfoField.Text = "Username:"
    $objInfoField.Location = New-Object System.Drawing.Size(10,8)
    $objTextBox = New-Object System.Windows.Forms.TextBox 
    $objTextBox.Location = New-Object System.Drawing.Size(10,27)
    $objTextBox.Size = New-Object System.Drawing.Size(130,20)
    $objTextBox.Focus()
    $objTextbox.Add_KeyDown({if ($_.KeyCode -eq "Enter")
        {FindADUser $objTextBox.Text}
    })
    $objSelectedUser2Label = New-Object System.Windows.Forms.Label
    $objSelectedUser2Label.Location = New-Object System.Drawing.Size(200,0)
    $objSelectedUser2Label.Size = New-Object System.Drawing.Size(90,15)
    $objSelectedUser2Label.Text = "User Selected:"
    $objSelectedUserLabel = New-Object System.Windows.Forms.Label
    $objSelectedUserLabel.Location = New-Object System.Drawing.Size(200,20)
    $objSelectedUserLabel.Size = New-Object System.Drawing.Size(210,20)
    $objSelectedUserLabel.Text = "NONE"
    $objCreationDateLabel = New-Object System.Windows.Forms.Label
    $objCreationDateLabel.Location = New-Object System.Drawing.Size(200,62)
    $objCreationDateLabel.Size = New-Object System.Drawing.Size(120,20)
    $objCreationDateLabel.Text = "Account Creation Date"
    $objCreationDateTextBox = New-Object System.Windows.Forms.TextBox
    $objCreationDateTextBox.Location = New-Object System.Drawing.Size(320,60)
    $objCreationDateTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objCreationDateTextBox.Text
    $objAccountStatusLabel = New-Object System.Windows.Forms.Label
    $objAccountStatusLabel.Location = New-Object System.Drawing.Size(200,92)
    $objAccountStatusLabel.AutoSize = $True
    $objAccountStatusLabel.Text = "Account Enabled"
    $objAccountStatusTextBox = New-Object System.Windows.Forms.TextBox
    $objAccountStatusTextBox.Location = New-Object System.Drawing.Size(320,90)
    $objAccountStatusTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objAccountStatusTextBox.Text
    $objMyIdentityLabel = New-Object System.Windows.Forms.Label 
    $objMyIdentityLabel.Location = New-Object System.Drawing.Size(200,122)
    $objMyIdentityLabel.AutoSize = $True
    $objMyIdentityLabel.Text = "MyIdentity Claimed"
    $objMyIdentityTextBox = New-Object System.Windows.Forms.TextBox
    $objMyIdentityTextBox.Location = New-Object System.Drawing.Size(320,120)
    $objMyIdentityTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objMyIdentityTextBox.Text
    $objDisplayNameLabel = New-Object System.Windows.Forms.Label
    $objDisplayNameLabel.Location = New-Object System.Drawing.Size(200,152)
    $objDisplayNameLabel.AutoSize = $True
    $objDisplayNameLabel.Text = "Display Name"
    $objDisplayNameTextbox = New-Object System.Windows.Forms.TextBox
    $objDisplayNameTextBox.Location = New-Object System.Drawing.Size(320,150)
    $objDisplayNameTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objDisplayNameTextBox.Text
    $objDistinguishedNameLabel = New-Object System.Windows.Forms.Label
    $objDistinguishedNameLabel.Location = New-Object System.Drawing.Size(200,182)
    $objDistinguishedNameLabel.AutoSize = $True
    $objDistinguishedNameLabel.Text = "Distinguished Name"
    $objDistinguishedNameTextBox = New-Object System.Windows.Forms.TextBox
    $objDistinguishedNameTextBox.Location = New-Object System.Drawing.Size(320,180)
    $objDistinguishedNameTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objDistinguishedNameTextBox.Text
    $objEmailLabel = New-Object System.Windows.Forms.Label
    $objEmailLabel.Location = New-Object System.Drawing.Size(200,212)
    $objEmailLabel.AutoSize = $True
    $objEmailLabel.Text = "Email"
    $objEmailTextBox = New-Object System.Windows.Forms.TextBox
    $objEmailTextBox.Location = New-Object System.Drawing.Size(320,210)
    $objEmailTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objEmailTextBox.Text
    $objEnterPriseIDLabel = New-Object System.Windows.Forms.Label
    $objEnterPriseIDLabel.Location = New-Object System.Drawing.Size(200,242)
    $objEnterPriseIDLabel.AutoSize = $True
    $objEnterPriseIDLabel.Text = "Enterprise ID"
    $objEnterPriseIDTextBox = New-Object System.Windows.Forms.TextBox
    $objEnterPriseIDTextBox.Location = New-Object System.Drawing.Size(320,240)
    $objEnterPriseIDTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objEnterPriseIDTextBox.Text
    $objJobTitleLabel = New-Object System.Windows.Forms.Label
    $objJobTitleLabel.Location = New-Object System.Drawing.Size(200,272)
    $objJobTitleLabel.AutoSize = $True
    $objJobTitleLabel.Text = "Job Title"
    $objJobTitleTextBox = New-Object System.Windows.Forms.TextBox
    $objJobTitleTextBox.Location = New-Object System.Drawing.Size(320,270)
    $objJobTitleTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objJobTitleTextBox.Text
    $objLastBadPasswordLabel = New-Object System.Windows.Forms.Label
    $objLastBadPasswordLabel.Location = New-Object System.Drawing.Size(200,302)
    $objLastBadPasswordLabel.AutoSize = $True
    $objLastBadPasswordLabel.Text = "Last Bad Password"
    $objLastBadPasswordTextBox = New-Object System.Windows.Forms.TextBox
    $objLastBadPasswordTextBox.Location = New-Object System.Drawing.Size(320,300)
    $objLastBadPasswordTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objLastBadPasswordTextBox.Text
    $objLastLogonLabel = New-Object System.Windows.Forms.Label
    $objLastLogonLabel.Location = New-Object System.Drawing.Size(200,332)
    $objLastLogonLabel.AutoSize = $True
    $objLastLogonLabel.Text = "Last Logon Time"
    $objLastLogonTextBox = New-Object System.Windows.Forms.TextBox
    $objLastLogonTextBox.Location = New-Object System.Drawing.Size(320,330)
    $objLastLogonTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objLastLogonTextBox.Text
    $objOrganizationalUnitLabel = New-Object System.Windows.Forms.Label
    $objOrganizationalUnitLabel.Location = New-Object System.Drawing.Size(200,362)
    $objOrganizationalUnitLabel.AutoSize = $True
    $objOrganizationalUnitLabel.Text = "Organizational Unit"
    $objOrganizationalUnitTextBox = New-Object System.Windows.Forms.TextBox
    $objOrganizationalUnitTextBox.Location = New-Object System.Drawing.Size(320,360)
    $objOrganizationalUnitTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objOrganizationalUnitTextBox.Text
    $objOffice365Label = New-Object System.Windows.Forms.Label
    $objOffice365Label.Location = New-Object System.Drawing.Size(200,392)
    $objOffice365Label.AutoSize = $True
    $objOffice365Label.Text = "Office 365 User"
    $objOffice365TextBox = New-Object System.Windows.Forms.TextBox
    $objOffice365TextBox.Location = New-Object System.Drawing.Size(320,392)
    $objOffice365TextBox.Size = New-Object System.Drawing.Size(200,20)
    $objOffice365TextBox.Text
    $objMobileEmailLabel = New-Object System.Windows.Forms.Label
    $objMobileEmailLabel.Location = New-Object System.Drawing.Size(200,422)
    $objMobileEmailLabel.AutoSize = $True
    $objMobileEmailLabel.Text = "Mobile Email"
    $objMobileEmailTextBox = New-Object System.Windows.Forms.TextBox
    $objMobileEmailTextBox.Location = New-Object System.Drawing.Size(320,420)
    $objMobileEmailTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objMobileEmailTextBox.Text
    $objWebClientLabel = New-Object System.Windows.Forms.Label
    $objWebClientLabel.Location = New-Object System.Drawing.Size(200,452)
    $objWebClientLabel.AutoSize = $True
    $objWebClientLabel.Text = "Outlook Client"
    $objWebClientTextBox = New-Object System.Windows.Forms.TextBox
    $objWebClientTextBox.Location = New-Object System.Drawing.Size(320,450)
    $objWebClientTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objWebClientTextBox.Text
    $objPasswordLastSetLabel = New-Object System.Windows.Forms.Label
    $objPasswordLastSetLabel.Location = New-Object System.Drawing.Size(200,482)
    $objPasswordLastSetLabel.AutoSize = $True
    $objPasswordLastSetLabel.Text = "Password Last Set"
    $objPasswordLastSetTextBox = New-Object System.Windows.Forms.TextBox
    $objPasswordLastSetTextBox.Location = New-Object System.Drawing.Size(320,480)
    $objPasswordLastSetTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objPasswordLastSetTextBox.Text
    $objPasswordStatusLabel = New-Object System.Windows.Forms.Label
    $objPasswordStatusLabel.Location = New-Object System.Drawing.Size(200,512)
    $objPasswordStatusLabel.AutoSize = $True
    $objPasswordStatusLabel.Text = "Password Expired"
    $objPasswordStatusTextBox = New-Object System.Windows.Forms.TextBox
    $objPasswordStatusTextBox.Location = New-Object System.Drawing.Size(320,510)
    $objPasswordStatusTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objPasswordStatusTextBox.Text
    $objPersonalDriveLabel = New-Object System.Windows.Forms.Label
    $objPersonalDriveLabel.Location = New-Object System.Drawing.Size(200,542)
    $objPersonalDriveLabel.AutoSize = $True
    $objPersonalDriveLabel.Text = "U Drive"
    $objPersonalDriveTextBox = New-Object System.Windows.Forms.TextBox
    $objPersonalDriveTextBox.Location = New-Object System.Drawing.Size(320,540)
    $objPersonalDriveTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objPersonalDriveTextBox.Text
    $objVPNUserLabel = New-Object System.Windows.Forms.Label
    $objVPNUserLabel.Location = New-Object System.Drawing.Size(200,572)
    $objVPNUserLabel.AutoSize = $True
    $objVPNUserLabel.Text = "VPN User"
    $objVPNUserTextBox = New-Object System.Windows.Forms.TextBox 
    $objVPNUserTextBox.Location = New-Object System.Drawing.Size(320,570)
    $objVPNUserTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objVPNUserTextBox.Text
    $objMemberOfLabel = New-Object System.Windows.Forms.Label
    $objMemberOfLabel.Location = New-Object System.Drawing.Size(10,100)
    $objMemberOfLabel.AutoSize = $True
    $objMemberOfLabel.Text = "Member Of:"
    $objMemberOfTextBox = New-Object System.Windows.Forms.ListBox
    $objMemberOfTextBox.Location = New-Object System.Drawing.Size(10,120)
    $objMemberOfTextBox.Size = New-Object System.Drawing.Size(175,400)
    $objChimeraForm.Controls.Add($objFindButton)
    $objChimeraForm.Controls.Add($objInfoField)
    $objChimeraForm.Controls.Add($objTextBox)
    $objChimeraForm.Controls.Add($objSelectedUser2Label)
    $objChimeraForm.Controls.Add($objSelectedUserLabel)
    $objChimeraForm.Controls.Add($objCreationDateLabel)
    $objChimeraForm.Controls.Add($objCreationDateTextBox)
    $objChimeraForm.Controls.Add($objAccountStatusLabel)
    $objChimeraForm.Controls.Add($objAccountStatusTextBox)
    $objChimeraForm.Controls.Add($objMyIdentityLabel)
    $objChimeraForm.Controls.Add($objMyIdentityTextBox)
    $objChimeraForm.Controls.Add($objDisplayNameLabel)
    $objChimeraForm.Controls.Add($objDisplayNameTextBox)
    $objChimeraForm.Controls.Add($objDistinguishedNameLabel)
    $objChimeraForm.Controls.Add($objDistinguishedNameTextBox)
    $objChimeraForm.Controls.Add($objEmailLabel)
    $objChimeraForm.Controls.Add($objEmailTextBox)
    $objChimeraForm.Controls.Add($objEnterPriseIDLabel)
    $objChimeraForm.Controls.Add($objEnterPriseIDTextBox)
    $objChimeraForm.Controls.Add($objJobTitleLabel)
    $objChimeraForm.Controls.Add($objJobTitleTextBox)
    $objChimeraForm.Controls.Add($objLastBadPasswordLabel)
    $objChimeraForm.Controls.Add($objLastBadPasswordTextBox)
    $objChimeraForm.Controls.Add($objLastLogonLabel)
    $objChimeraForm.Controls.Add($objLastLogonTextBox)
    $objChimeraForm.Controls.Add($objOrganizationalUnitLabel)
    $objChimeraForm.Controls.Add($objOrganizationalUnitTextBox)
    $objChimeraForm.Controls.Add($objOffice365Label)
    $objChimeraForm.Controls.Add($objOffice365TextBox)
    $objChimeraForm.Controls.Add($objMobileEmailLabel)
    $objChimeraForm.Controls.Add($objMobileEmailTextBox)
    $objChimeraForm.Controls.Add($objWebClientLabel)
    $objChimeraForm.Controls.Add($objWebClientTextBox)
    $objChimeraForm.Controls.Add($objPasswordLastSetLabel)
    $objChimeraForm.Controls.Add($objPasswordLastSetTextBox)
    $objChimeraForm.Controls.Add($objPasswordStatusLabel)
    $objChimeraForm.Controls.Add($objPasswordStatusTextBox)
    $objChimeraForm.Controls.Add($objPersonalDriveLabel)
    $objChimeraForm.Controls.Add($objPersonalDriveTextBox)
    $objChimeraForm.Controls.Add($objVPNUserLabel)
    $objChimeraForm.Controls.Add($objVPNUserTextBox)
    $objChimeraForm.Controls.Add($objMemberOfLabel)
    $objChimeraForm.Controls.Add($objMemberOfTextBox)
    $objFunctionsLabel = New-Object System.Windows.Forms.Label
    $objFunctionsLabel.Location = New-Object System.Drawing.Size(575,8)
    $objFunctionsLabel.AutoSize = $True
    $objFunctionsLabel.Text = "Computer Tag:"
    $objPCTagTextBox = New-Object System.Windows.Forms.TextBox
    $objPCTagTextBox.Location = New-Object System.Drawing.Size(575,27)
    $objPCTagTextBox.Size = New-Object System.Drawing.Size(200,20)
    $objPCTagTextBox.Add_KeyDown({ if ( $_.KeyCode -eq "Enter" ) {
        [Windows.Forms.MessageBox]::Show("Please Select an Option Below", "No Option Selected", [Windows.Forms.MessageBoxButtons]::OK, [Windows.Forms.MessageBoxIcon]::Error)
    }})
    $objFunction0Button = New-Object System.Windows.Forms.Button
    $objFunction0Button.Location = New-Object System.Drawing.Size(575,55)
    $objFunction0Button.Size = New-Object System.Drawing.Size(200,23)
    $objFunction0Button.Text = "Computer File Explorer"
    $objFunction0Button.Add_Click({ComputerFileExplorer ($objPCTagTextBox)})
    $objFunction1Button = New-Object System.Windows.Forms.Button
    $objFunction1Button.Location = New-Object System.Drawing.Size(575,85)
    $objFunction1Button.Size = New-Object System.Drawing.Size(200,23)
    $objFunction1Button.Text = "Laps Password Viewer"
    $objFunction1Button.Add_Click({LAPS ($objPCTagTextBox)})
    $objFunction2Button = New-Object System.Windows.Forms.Button
    $objFunction2Button.Location = New-Object System.Drawing.Size(575,115)
    $objFunction2Button.Size = New-Object System.Drawing.Size(200,23)
    $objFunction2Button.Text = "Remote to Computer"
    $objFunction2Button.Add_Click({REMOTE ($objPCTagTextBox)})
    $objFunction3Button = New-Object System.Windows.Forms.Button
    $objFunction3Button.Location = New-Object System.Drawing.Size(575,145)
    $objFunction3Button.Size = New-Object System.Drawing.Size(200,23)
    $objFunction3Button.Text = "Reboot Computer w/ Auto Remote in"
    $objFunction3Button.Add_Click({REBOOTREMOTE ($objPCTagTextBox)})
    $objFunction4Button = New-Object System.Windows.Forms.Button
    $objFunction4Button.Location = New-Object System.Drawing.Size(575,175)
    $objFunction4Button.Size = New-Object System.Drawing.Size(200,23)
    $objFunction4Button.Text = "Windows/IGEL Check"
    $objFunction4Button.Add_Click({WINDOWSIGELCHECK ($objPCTagTextBox)})
    $objFunction5Button = New-Object System.Windows.Forms.Button
    $objFunction5Button.Location = New-Object System.Drawing.Size(575,205)
    $objFunction5Button.Size = New-Object System.Drawing.Size(200,23)
    $objFunction5Button.Text = "Remote Desktop User Modification"
    $objFunction5Button.Add_Click({REMOTEDESKTOPUSER ($objPCTagTextBox)})
    $objFunctionHelpButton = New-Object System.Windows.Forms.Button
    $objFunctionHelpButton.Location = New-Object System.Drawing.Size(10,525)
    $objFunctionHelpButton.Size = New-Object System.Drawing.Size(73,23)
    $objFunctionHelpButton.Text = "Help"
    $objFunctionHelpButton.Add_Click({HELP})
    $objFunctionWhatsNewButton = New-Object System.Windows.Forms.Button
    $objFunctionWhatsNewButton.Location = New-Object System.Drawing.Size(10,555)
    $objFunctionWhatsNewButton.Size = New-Object System.Drawing.Size(73,23)
    $objFunctionWhatsNewButton.Text = "Whats New"
    $objFunctionWhatsNewButton.Add_Click({WHATSNEW})
    $objWarningMessage = New-Object System.Windows.Forms.Label
    $objWarningMessage.Text = "WARNING: This script is no longer maintained."
    $objWarningMessage.Location = New-Object System.Drawing.Size(575,235)
    $objWarningMessage.Size = New-Object System.Drawing.Size(200,75)
    $objWarningMessage.ForeColor = 'Red'
    $objChimeraForm.Controls.Add($objWarningMessage)
    $objChimeraForm.Controls.Add($objFunctionsLabel)
    $objChimeraForm.Controls.Add($objPCTagTextBox)
    $objChimeraForm.Controls.Add($objFunction0Button)
    $objChimeraForm.Controls.Add($objFunction1Button)
    $objChimeraForm.Controls.Add($objFunction2Button)
    $objChimeraForm.Controls.Add($objFunction3Button)
    $objChimeraForm.Controls.Add($objFunction4Button)
    $objChimeraForm.Controls.Add($objFunction5Button)
    $objChimeraForm.Controls.Add($objFunctionHelpButton)
    $objChimeraForm.Controls.Add($objFunctionWhatsNewButton)
$objChimeraForm.Add_Shown({$objChimeraForm.Activate(); $objTextBox.Focus()})
[void]$objChimeraForm.ShowDialog()
