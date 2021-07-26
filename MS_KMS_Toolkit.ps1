#
# Author: Lukas Reimann
# 
# License:
#
#    This program is free software: you can redistribute it and/or modify
#    it under the terms of the GNU General Public License as published by
#    the Free Software Foundation, either version 3 of the License, or
#    (at your option) any later version.
#
#    This program is distributed in the hope that it will be useful,
#    but WITHOUT ANY WARRANTY; without even the implied warranty of
#    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#    GNU General Public License for more details.
#
#    You should have received a copy of the GNU General Public License
#    along with this program.  If not, see <https://www.gnu.org/licenses/>.
#
#


#region clear syntax mirror
    cls 
#endregion

#region add assambly
    Add-Type -AssemblyName System.Windows.Forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
#endregion

#region Changelog
    #
    # 1.0.1: First published version. 
    # 1.0.2: Added Client licensing status check, Added client remote KMS activation function
    # 1.0.3: Added check installed windows version function
    # 
#endregion

#region Variables
    
    #------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    #   Variable                       | Value                                                                                                  |Description
    #------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    $Version =                         "1.0.3"                                                                                                    # Published script version
    #------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

#endregion

#region GUI Element definitions

    #region Windows Form

        $form                            = New-Object system.Windows.Forms.Form
        $form.ClientSize                 = '1500,1000'
        $form.text                       = "KMS Management Toolkit"
        $form.TopMost                    = $false
        $form.BackColor                  = "#FFFFFF"
        $form.AutoScroll                 = $true
        $form.MinimizeBox                = $false
        $form.MaximizeBox                = $false
        $form.ControlBox                 = $false

    #endregion

    #region Label 'Version'
        $l_Version                         = New-Object system.Windows.Forms.Label
        $l_Version.text                    = "Version: $Version"
        $l_Version.AutoSize                = $true
        $l_Version.width                   = 100
        $l_Version.height                  = 30
        $l_Version.location                = New-Object System.Drawing.Point(1350,50)
        $l_Version.Font                    = 'Microsoft Sans Serif,15'
        $l_Version.visible                 = $true
        $l_Version.ForeColor               = "#000000"
        $l_Version.BackColor               = "#ffffff"

        $form.Controls.Add($l_Version)
    #endregion

    #region Button 'Exit'
        $b_Exit                         = New-Object system.Windows.Forms.Button
        $b_Exit.BackColor               = "#FFFFFF"
        $b_Exit.text                    = "Exit"
        $b_Exit.width                   = 100
        $b_Exit.height                  = 30
        $b_Exit.location                = New-Object System.Drawing.Point(1380,10)
        $b_Exit.Font                    = 'Microsoft Sans Serif,10'#
        $b_Exit.Visible                 = $true

        $form.Controls.Add($b_Exit)

        $b_Exit.Add_Click({
            $form.Close()
        })
    #endregion

    #region Button 'Refresh'
        $b_Refresh                         = New-Object system.Windows.Forms.Button
        $b_Refresh.BackColor               = "#FFFFFF"
        $b_Refresh.text                    = "Refresh"
        $b_Refresh.width                   = 200
        $b_Refresh.height                  = 30
        $b_Refresh.location                = New-Object System.Drawing.Point(811,930)
        $b_Refresh.Font                    = 'Microsoft Sans Serif,10'#
        $b_Refresh.Visible                 = $true

        $form.Controls.Add($b_Refresh)

        $b_Refresh.Add_Click({
            func-FillListCurrentLicensedClients
        })
    #endregion

    #region Label 'Current licensed clients'
        $l_CurrentLicensedClients                         = New-Object system.Windows.Forms.Label
        $l_CurrentLicensedClients.text                    = "Current Licensed Clients"
        $l_CurrentLicensedClients.AutoSize                = $true
        $l_CurrentLicensedClients.width                   = 100
        $l_CurrentLicensedClients.height                  = 30
        $l_CurrentLicensedClients.location                = New-Object System.Drawing.Point(50,80)
        $l_CurrentLicensedClients.Font                    = 'Microsoft Sans Serif,15'
        $l_CurrentLicensedClients.visible                 = $true
        $l_CurrentLicensedClients.ForeColor               = "#000000"
        $l_CurrentLicensedClients.BackColor               = "#ffffff"

        $form.Controls.Add($l_CurrentLicensedClients)
    #endregion

    #region Textbox current KMS clients

        $t_CurrentLicensedClients = New-Object System.Windows.Forms.TextBox
        $t_CurrentLicensedClients.Location = New-Object System.Drawing.Point(10,120)
        $t_CurrentLicensedClients.Size = New-Object System.Drawing.Size(1000,800)
        $t_CurrentLicensedClients.Scrollbars = "Vertical" 
        $t_CurrentLicensedClients.Multiline = $True;
        $t_CurrentLicensedClients.Font = "Microsoft Sans Serif,15"

        $form.Controls.Add($t_CurrentLicensedClients)

    #endregion

    #region "Client License Status Check"

        #region Label 'Client License Check'
            $l_ClientCheckTitle                         = New-Object system.Windows.Forms.Label
            $l_ClientCheckTitle.text                    = "Client License Status Check"
            $l_ClientCheckTitle.AutoSize                = $true
            $l_ClientCheckTitle.width                   = 100
            $l_ClientCheckTitle.height                  = 30
            $l_ClientCheckTitle.location                = New-Object System.Drawing.Point(1100,120)
            $l_ClientCheckTitle.Font                    = 'Microsoft Sans Serif,15'
            $l_ClientCheckTitle.visible                 = $true
            $l_ClientCheckTitle.ForeColor               = "#000000"
            $l_ClientCheckTitle.BackColor               = "#ffffff"

            $form.Controls.Add($l_ClientCheckTitle)
        #endregion

        #region Label 'Client License Check'
            $l_ClientCheckerInfo                         = New-Object system.Windows.Forms.Label
            $l_ClientCheckerInfo.text                    = "Enter FQDN of a client you want to check"
            $l_ClientCheckerInfo.AutoSize                = $true
            $l_ClientCheckerInfo.width                   = 100
            $l_ClientCheckerInfo.height                  = 30
            $l_ClientCheckerInfo.location                = New-Object System.Drawing.Point(1050,190)
            $l_ClientCheckerInfo.Font                    = 'Microsoft Sans Serif,10'
            $l_ClientCheckerInfo.visible                 = $true
            $l_ClientCheckerInfo.ForeColor               = "#000000"
            $l_ClientCheckerInfo.BackColor               = "#ffffff"

            $form.Controls.Add($l_ClientCheckerInfo)
        #endregion

        #region Textbox current KMS clients

            $t_ClientToCheck = New-Object System.Windows.Forms.TextBox
            $t_ClientToCheck.Location = New-Object System.Drawing.Point(1050,170)
            $t_ClientToCheck.Size = New-Object System.Drawing.Size(300,30)
            $t_ClientToCheck.Multiline = $false

            $form.Controls.Add($t_ClientToCheck)

        #endregion

        #region Button 'Refresh'
            $b_CheckClient                         = New-Object system.Windows.Forms.Button
            $b_CheckClient.BackColor               = "#FFFFFF"
            $b_CheckClient.text                    = "Check Client"
            $b_CheckClient.width                   = 100
            $b_CheckClient.height                  = 30
            $b_CheckClient.location                = New-Object System.Drawing.Point(1370,165)
            $b_CheckClient.Font                    = 'Microsoft Sans Serif,10'#
            $b_CheckClient.Visible                 = $true

            $form.Controls.Add($b_CheckClient)

            $b_CheckClient.Add_Click({
                func-CheckClientStatus
            })
        #endregion

        #region Status: 0 Checkbox
            
            # Windows activation state value = 0 

            $t_Status0 = New-Object System.Windows.Forms.TextBox
            $t_Status0.Location = New-Object System.Drawing.Point(1050,220)
            $t_Status0.Size = New-Object System.Drawing.Size(20,20)
            $t_Status0.Multiline = $false
            $t_Status0.Text = ""

            $form.Controls.Add($t_Status0)

        #endregion

        #region Status: 0 Label

            # Windows activation state value = 0

            $l_Status0                         = New-Object system.Windows.Forms.Label
            $l_Status0.text                    = "Unlicensed"
            $l_Status0.AutoSize                = $true
            $l_Status0.width                   = 100
            $l_Status0.height                  = 30
            $l_Status0.location                = New-Object System.Drawing.Point(1080,222)
            $l_Status0.Font                    = 'Microsoft Sans Serif,10'
            $l_Status0.visible                 = $true
            $l_Status0.ForeColor               = "#000000"
            $l_Status0.BackColor               = "#ffffff"

            $form.Controls.Add($l_Status0)

        #endregion

        #region Status: 1 Checkbox

            # Windows activation state value = 1

            $t_Status1 = New-Object System.Windows.Forms.TextBox
            $t_Status1.Location = New-Object System.Drawing.Point(1200,220)
            $t_Status1.Size = New-Object System.Drawing.Size(20,20)
            $t_Status1.Multiline = $false
            $t_Status1.Text = ""

            $form.Controls.Add($t_Status1)

        #endregion

        #region Status: 1 Label

            # Windows activation state value = 1

            $l_Status1                         = New-Object system.Windows.Forms.Label
            $l_Status1.text                    = "Licensed"
            $l_Status1.AutoSize                = $true
            $l_Status1.width                   = 100
            $l_Status1.height                  = 30
            $l_Status1.location                = New-Object System.Drawing.Point(1230,222)
            $l_Status1.Font                    = 'Microsoft Sans Serif,10'
            $l_Status1.visible                 = $true
            $l_Status1.ForeColor               = "#000000"
            $l_Status1.BackColor               = "#ffffff"

            $form.Controls.Add($l_Status1)

        #endregion

        #region Status: 2 Checkbox

            # Windows activation state value = 2

            $t_Status2 = New-Object System.Windows.Forms.TextBox
            $t_Status2.Location = New-Object System.Drawing.Point(1050,250)
            $t_Status2.Size = New-Object System.Drawing.Size(20,20)
            $t_Status2.Multiline = $false
            $t_Status2.Text = ""

            $form.Controls.Add($t_Status2)

        #endregion

        #region Status: 2 Label

            # Windows activation state value = 2

            $l_Status2                         = New-Object system.Windows.Forms.Label
            $l_Status2.text                    = "OOBGrace"
            $l_Status2.AutoSize                = $true
            $l_Status2.width                   = 100
            $l_Status2.height                  = 30
            $l_Status2.location                = New-Object System.Drawing.Point(1080,252)
            $l_Status2.Font                    = 'Microsoft Sans Serif,10'
            $l_Status2.visible                 = $true
            $l_Status2.ForeColor               = "#000000"
            $l_Status2.BackColor               = "#ffffff"

            $form.Controls.Add($l_Status2)

        #endregion

        #region Status: 3 Checkbox

            # Windows activation state value = 3

            $t_Status3 = New-Object System.Windows.Forms.TextBox
            $t_Status3.Location = New-Object System.Drawing.Point(1200,250)
            $t_Status3.Size = New-Object System.Drawing.Size(20,20)
            $t_Status3.Multiline = $false
            $t_Status3.Text = ""

            $form.Controls.Add($t_Status3)

        #endregion

        #region Status: 3 Label

            # Windows activation state value = 3

            $l_Status3                         = New-Object system.Windows.Forms.Label
            $l_Status3.text                    = "OOTGrace"
            $l_Status3.AutoSize                = $true
            $l_Status3.width                   = 100
            $l_Status3.height                  = 30
            $l_Status3.location                = New-Object System.Drawing.Point(1230,252)
            $l_Status3.Font                    = 'Microsoft Sans Serif,10'
            $l_Status3.visible                 = $true
            $l_Status3.ForeColor               = "#000000"
            $l_Status3.BackColor               = "#ffffff"

            $form.Controls.Add($l_Status3)

        #endregion

        #region Status: 4 Checkbox

            # Windows activation state value = 4

            $t_Status4 = New-Object System.Windows.Forms.TextBox
            $t_Status4.Location = New-Object System.Drawing.Point(1050,280)
            $t_Status4.Size = New-Object System.Drawing.Size(20,20)
            $t_Status4.Multiline = $false
            $t_Status4.Text = ""

            $form.Controls.Add($t_Status4)

        #endregion

        #region Status: 4 Label
            
            # Windows activation state value = 4

            $l_Status4                         = New-Object system.Windows.Forms.Label
            $l_Status4.text                    = "NonGenuineGrace"
            $l_Status4.AutoSize                = $true
            $l_Status4.width                   = 100
            $l_Status4.height                  = 30
            $l_Status4.location                = New-Object System.Drawing.Point(1080,282)
            $l_Status4.Font                    = 'Microsoft Sans Serif,10'
            $l_Status4.visible                 = $true
            $l_Status4.ForeColor               = "#000000"
            $l_Status4.BackColor               = "#ffffff"

            $form.Controls.Add($l_Status4)

        #endregion

        #region Status: 5 Checkbox
            
            # Windows activation state value = 5

            $t_Status5 = New-Object System.Windows.Forms.TextBox
            $t_Status5.Location = New-Object System.Drawing.Point(1200,280)
            $t_Status5.Size = New-Object System.Drawing.Size(20,20)
            $t_Status5.Multiline = $false
            $t_Status5.Text = ""

            $form.Controls.Add($t_Status5)

        #endregion

        #region Status: 5 Label
            
            # Windows activation state value = 5

            $l_Status5                         = New-Object system.Windows.Forms.Label
            $l_Status5.text                    = "Notification"
            $l_Status5.AutoSize                = $true
            $l_Status5.width                   = 100
            $l_Status5.height                  = 30
            $l_Status5.location                = New-Object System.Drawing.Point(1230,282)
            $l_Status5.Font                    = 'Microsoft Sans Serif,10'
            $l_Status5.visible                 = $true
            $l_Status5.ForeColor               = "#000000"
            $l_Status5.BackColor               = "#ffffff"

            $form.Controls.Add($l_Status5)

        #endregion

        #region Status: 6 Checkbox
            
            # Windows activation state value = 5

            $t_Status6 = New-Object System.Windows.Forms.TextBox
            $t_Status6.Location = New-Object System.Drawing.Point(1200,310)
            $t_Status6.Size = New-Object System.Drawing.Size(20,20)
            $t_Status6.Multiline = $false
            $t_Status6.Text = ""

            $form.Controls.Add($t_Status6)

        #endregion

        #region Status: 6 Label
            
            # Windows activation state value = 6

            $l_Status6                         = New-Object system.Windows.Forms.Label
            $l_Status6.text                    = "ExtendedGrace"
            $l_Status6.AutoSize                = $true
            $l_Status6.width                   = 100
            $l_Status6.height                  = 30
            $l_Status6.location                = New-Object System.Drawing.Point(1230,312)
            $l_Status6.Font                    = 'Microsoft Sans Serif,10'
            $l_Status6.visible                 = $true
            $l_Status6.ForeColor               = "#000000"
            $l_Status6.BackColor               = "#ffffff"

            $form.Controls.Add($l_Status6)

        #endregion

        #region Client OS
            
            # Windows activation state value = 5

            $l_ClientOS                         = New-Object system.Windows.Forms.Label
            $l_ClientOS.text                    = "-"
            $l_ClientOS.AutoSize                = $true
            $l_ClientOS.width                   = 100
            $l_ClientOS.height                  = 30
            $l_ClientOS.location                = New-Object System.Drawing.Point(1050,350)
            $l_ClientOS.Font                    = 'Microsoft Sans Serif,10'
            $l_ClientOS.visible                 = $true
            $l_ClientOS.ForeColor               = "#000000"
            $l_ClientOS.BackColor               = "#ffffff"

            $form.Controls.Add($l_ClientOS)

        #endregion
        

    #endregion

    #region "Install License Key"

        #region Label 'Install and Activate License Key'
            $l_IPKTitle                         = New-Object system.Windows.Forms.Label
            $l_IPKTitle.text                    = "Install and Activate License Key"
            $l_IPKTitle.AutoSize                = $true
            $l_IPKTitle.width                   = 100
            $l_IPKTitle.height                  = 30
            $l_IPKTitle.location                = New-Object System.Drawing.Point(1100,400)
            $l_IPKTitle.Font                    = 'Microsoft Sans Serif,15'
            $l_IPKTitle.visible                 = $true
            $l_IPKTitle.ForeColor               = "#000000"
            $l_IPKTitle.BackColor               = "#ffffff"

            $form.Controls.Add($l_IPKTitle)
        #endregion

        #region KMS Key: Windows Server 2016 Data Center
            $r_KMS1 = New-Object System.Windows.Forms.RadioButton
            $r_KMS1.Location = '1050,450'
            $r_KMS1.size = '200,20'
            $r_KMS1.Checked = $false 
            $r_KMS1.Text = "Windows Server 2016 Data Center"

            $form.Controls.Add($r_KMS1)
        #endregion

        #region KMS Key: Windows Server 2016 Standard
            $r_KMS2 = New-Object System.Windows.Forms.RadioButton
            $r_KMS2.Location = '1050,480'
            $r_KMS2.size = '200,20'
            $r_KMS2.Checked = $false
            $r_KMS2.Text = "Windows Server 2016 Standard"

            $form.Controls.Add($r_KMS2)
        #endregion

        #region KMS Key: Windows Server 2016 Essential
            $r_KMS3 = New-Object System.Windows.Forms.RadioButton
            $r_KMS3.Location = '1050,510'
            $r_KMS3.size = '200,20'
            $r_KMS3.Checked = $false
            $r_KMS3.Text = "Windows Server 2016 Essential"

            $form.Controls.Add($r_KMS3)
        #endregion

        #region KMS Key: Windows Server 2019 Data Center
            $r_KMS4 = New-Object System.Windows.Forms.RadioButton
            $r_KMS4.Location = '1300,450'
            $r_KMS4.size = '200,20'
            $r_KMS4.Checked = $false 
            $r_KMS4.Text = "Windows Server 2019 Data Center"

            $form.Controls.Add($r_KMS4)
        #endregion

        #region KMS Key: Windows Server 2019 Standard
            $r_KMS5 = New-Object System.Windows.Forms.RadioButton
            $r_KMS5.Location = '1300,480'
            $r_KMS5.size = '200,20'
            $r_KMS5.Checked = $false
            $r_KMS5.Text = "Windows Server 2019 Standard"

            $form.Controls.Add($r_KMS5)
        #endregion

        #region KMS Key: Windows Server 2019 Essentia
            $r_KMS6 = New-Object System.Windows.Forms.RadioButton
            $r_KMS6.Location = '1300,510'
            $r_KMS6.size = '200,20'
            $r_KMS6.Checked = $false
            $r_KMS6.Text = "Windows Server 2019 Essential"

            $form.Controls.Add($r_KMS6)
        #endregion

        #region Textbox FQDN of the client you want to IPK

            $t_ClientToIPK = New-Object System.Windows.Forms.TextBox
            $t_ClientToIPK.Location = New-Object System.Drawing.Point(1050,550)
            $t_ClientToIPK.Size = New-Object System.Drawing.Size(300,30)
            $t_ClientToIPK.Multiline = $false

            $form.Controls.Add($t_ClientToIPK)

        #endregion

        #region Button 'Exit'
        $b_IPK                         = New-Object system.Windows.Forms.Button
        $b_IPK.BackColor               = "#FFFFFF"
        $b_IPK.text                    = "Remote IPK"
        $b_IPK.width                   = 100
        $b_IPK.height                  = 30
        $b_IPK.location                = New-Object System.Drawing.Point(1370,544)
        $b_IPK.Font                    = 'Microsoft Sans Serif,10'#
        $b_IPK.Visible                 = $true

        $form.Controls.Add($b_IPK)

        $b_IPK.Add_Click({
            func-IPK
        })
    #endregion

    #endregion

#endregion

function func-FillListCurrentLicensedClients{
    $t_CurrentLicensedClients.text = $(foreach ($entry in (Get-EventLog -Logname "Key Management Service")) {$entry.ReplacementStrings[3]}) | sort-object -Unique | Sort-Object name | Out-String
}

function func-CheckClientStatus{

    $t_Status0.Text = " "
    $t_Status1.Text = " "
    $t_Status2.Text = " "
    $t_Status3.Text = " "
    $t_Status4.Text = " "
    $t_Status5.Text = " "
    $t_Status6.Text = " "

    
    $RemoteComputer = $t_ClientToCheck.text

    Enter-PSSession -computername $RemoteComputer 

    $LICSTATEproperty = Get-CimInstance SoftwareLicensingProduct -ComputerName $RemoteComputer -Filter "Name like 'Windows%'" | where { $_.PartialProductKey } 

    $LICSTATE = $LICSTATEproperty.LicenseStatus

    #$CLIENTOS = $LICSTATEproperty.Description

    #$l_ClientOS.text = "$CLIENTOS"

    if($LICSTATE -eq "0"){
        Write-Host "Status 0"
        $t_Status0.Text = " X"
    }

    if($LICSTATE -eq "1"){
        Write-Host "Status 1"
        $t_Status1.Text = " X"
    }

    if($LICSTATE -eq "2"){
        Write-Host "Status 2"
        $t_Status2.Text = " X"
    }

    if($LICSTATE -eq "3"){
        Write-Host "Status 3"
        $t_Status3.Text = " X"
    }

    if($LICSTATE -eq "4"){
        Write-Host "Status 4"
        $t_Status4.Text = " X"
    }

    if($LICSTATE -eq "5"){
        Write-Host "Status 5"
        $t_Status5.Text = " X"
    }

    if($LICSTATE -eq "6"){
        Write-Host "Status 6"
        $t_Status6.Text = " X"
    }
  
    #exit remote PS Session
    Exit-PSSession

    func-CheckWindowsVersion
}

function func-CheckWindowsVersion{
    
    $WindowsVersion = Invoke-Command -ComputerName $RemoteComputer -ScriptBlock {(Get-WmiObject -class Win32_OperatingSystem).Caption} -Verbose

    $l_ClientOS.text = "$WindowsVersion"

    # (Get-WmiObject -class Win32_OperatingSystem).Caption
}

function func-IPK{

    $IPKComputer = $t_ClientToIPK.Text

    #KMS Setup Keys
    #Server 2022
    $KMS_2022_DC = "WX4NM-KYWYW-QJJR4-XV3QB-6VM33"
    $KMS_2022_STD = "VDYBN-27WPP-V4HQT-9VMD4-VMK7H"

    #Server 2019
    $KMS_2019_DC = "WMDGN-G9PQG-XVVXX-R3X43-63DFG"
    $KMS_2019_STD = "N69G4-B89J2-4G8F4-WWYCC-J464C"
    $KMS_2019_ES = "WVDHN-86M7X-466P6-VHXV7-YY726"

    #Server 2016
    $KMS_2016_DC = "CB7KF-BWN84-R7R2Y-793K2-8XDDG"
    $KMS_2016_STD = "WC2BQ-8NRM3-FDDYY-2BFGV-KHKQY"
    $KMS_2016_ES = "JCKRF-N37P4-C2D82-9YXRT-4M63B"

    if($r_kms1.Checked -eq $true){
        $RemoteActivationKey = $KMS_2016_DC
    }

    if($r_kms2.Checked -eq $true){
        $RemoteActivationKey = $KMS_2016_STD
    }

    if($r_kms3.Checked -eq $true){
        $RemoteActivationKey = $KMS_2016_ES
    }

    if($r_kms4.Checked -eq $true){
        $RemoteActivationKey = $KMS_2019_DC
    }

    if($r_kms5.Checked -eq $true){
        $RemoteActivationKey = $KMS_2019_STD
    }

    if($r_kms6.Checked -eq $true){
        $RemoteActivationKey = $KMS_2019_ES
    }
    
    #Start invoke command for running the scriptblock at the client you want to IPK remotely. The selected remote activation key will be transmitted from the local scope to the remote scope ($args[0])
    Invoke-Command -ComputerName $IPKComputer -ArgumentList $RemoteActivationKey -ScriptBlock {cscript.exe "$env:SystemRoot\System32\slmgr.vbs" /ipk $args[0]} -Verbose

}

func-FillListCurrentLicensedClients

#region form dialog
    $form.controls.AddRange(@())
    [void]$form.ShowDialog()
#endregion