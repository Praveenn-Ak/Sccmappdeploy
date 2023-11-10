<# Script inputs that required to be modified.
line # 86
Approver email Id's
$Approverlist.Items.AddRange([System.Object[]](@("huchaiah@abc_corp.com", "govindaraj@abc_corp.com", "P.km@abc_corp.com", "K.kumar@abc_corp.com", "V.mc@abc_corp.com", "G.m@abc_corp.com ")))

line 291-293
# Site configuration
$SiteCode = "ABC" # Site code 
$ProviderMachineName = "Server-25" # SMS Provider machine name

line 356
$SMTPServer = "App-mail.abc.corp.bce.ca"

#>

#region Script Settings

#<ScriptSettings>
#  <ScriptPackager>
#    <process>powershell.exe</process>
#    <arguments />
#    <extractdir>%TEMP%</extractdir>
#    <files />
#    <usedefaulticon>true</usedefaulticon>
#    <showinsystray>false</showinsystray>
#    <altcreds>false</altcreds>
#    <efs>true</efs>
#    <ntfs>true</ntfs>
#    <local>false</local>
#    <abortonfail>true</abortonfail>
#    <product />
#    <version>1.0.0.1</version>
#    <versionstring />
#    <comments />
#    <company />
#    <includeinterpreter>false</includeinterpreter>
#    <forcecomregistration>false</forcecomregistration>
#    <consolemode>false</consolemode>
#    <EnableChangelog>false</EnableChangelog>
#    <AutoBackup>false</AutoBackup>
#    <snapinforce>false</snapinforce>
#    <snapinshowprogress>false</snapinshowprogress>
#    <snapinautoadd>2</snapinautoadd>
#    <snapinpermanentpath />
#    <cpumode>1</cpumode>
#    <hidepsconsole>false</hidepsconsole>
#  </ScriptPackager>
#</ScriptSettings>
#endregion

#region ScriptForm Designer

#region Constructor

[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#endregion

#region Post-Constructor Custom Code

#endregion

#region Form Creation
#Warning: It is recommended that changes inside this region be handled using the ScriptForm Designer.
#When working with the ScriptForm designer this region and any changes within may be overwritten.
#~~< Form1 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Form1 = New-Object System.Windows.Forms.Form
$Form1.ClientSize = New-Object System.Drawing.Size(1050, 612)
$Form1.Text = "SCCM Application Deployment Tool"
#~~< CRQ_REQ_NO >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CRQ_REQ_NO = New-Object System.Windows.Forms.Label
$CRQ_REQ_NO.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$CRQ_REQ_NO.Font = New-Object System.Drawing.Font("Tahoma", 10.0, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$CRQ_REQ_NO.Location = New-Object System.Drawing.Point(62, 28)
$CRQ_REQ_NO.Size = New-Object System.Drawing.Size(176, 21)
$CRQ_REQ_NO.Text = "CRQ or REQ NO"
#~~< CRQ_REQ >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$CRQ_REQ = New-Object System.Windows.Forms.TextBox
$CRQ_REQ.Location = New-Object System.Drawing.Point(262, 28)
$CRQ_REQ.Size = New-Object System.Drawing.Size(290, 21)
#$CRQ_REQ.TabIndex = 1
$CRQ_REQ.Text = ""
$CRQ_REQ.BackColor = [System.Drawing.SystemColors]::Window
#~~< Approverlist >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Approverlist = New-Object System.Windows.Forms.ComboBox
$Approverlist.FormattingEnabled = $true
$Approverlist.Location = New-Object System.Drawing.Point(262, 73)
$Approverlist.Size = New-Object System.Drawing.Size(290, 21)
$Approverlist.TabIndex = 20
$Approverlist.Text = ""
$Approverlist.Items.AddRange([System.Object[]](@("huchaiah@abc_corp.com", "govindaraj@abc_corp.com", "P.km@abc_corp.com", "K.kumar@abc_corp.com", "V.mc@abc_corp.com", "G.m@abc_corp.com ")))
$Approverlist.SelectedIndex = -1
#~~< Approver >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Approver = New-Object System.Windows.Forms.Label
$Approver.AccessibleDescription = "SCCM Application Deployment Approver"
$Approver.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$Approver.Font = New-Object System.Drawing.Font("Tahoma", 10.0)
$Approver.Location = New-Object System.Drawing.Point(62, 73)
$Approver.Size = New-Object System.Drawing.Size(176, 21)
$Approver.TabIndex = 19
$Approver.Text = "Deployment Approver"
#~~< Log >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Log = New-Object System.Windows.Forms.RichTextBox
$Log.Font = New-Object System.Drawing.Font("Tahoma", 12.0, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$Log.Location = New-Object System.Drawing.Point(592, 28)
$Log.ReadOnly = $true
$Log.Size = New-Object System.Drawing.Size(425, 479)
$Log.TabIndex = 6
$Log.Text = ""
$Log.BackColor = [System.Drawing.SystemColors]::Window
$Log.ForeColor = [System.Drawing.SystemColors]::Desktop
#~~< RequiredUTC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$RequiredUTC = New-Object System.Windows.Forms.CheckBox
$RequiredUTC.Location = New-Object System.Drawing.Point(477, 392)
$RequiredUTC.Size = New-Object System.Drawing.Size(104, 24)
$RequiredUTC.TabIndex = 11
$RequiredUTC.Text = "UTC"
$RequiredUTC.UseVisualStyleBackColor = $true
#~~< AvailableUTC >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$AvailableUTC = New-Object System.Windows.Forms.CheckBox
$AvailableUTC.Location = New-Object System.Drawing.Point(477, 339)
$AvailableUTC.Size = New-Object System.Drawing.Size(104, 24)
$AvailableUTC.TabIndex = 10
$AvailableUTC.Text = "UTC"
$AvailableUTC.UseVisualStyleBackColor = $true
#~~< Required_DateTime >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Required_DateTime = New-Object System.Windows.Forms.DateTimePicker
$Required_DateTime.AllowDrop = $true
$Required_DateTime.CustomFormat = "yyyy-MM-dd hh:mm tt"
$Required_DateTime.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$Required_DateTime.Location = New-Object System.Drawing.Point(262, 392)
$Required_DateTime.Size = New-Object System.Drawing.Size(187, 20)
$Required_DateTime.TabIndex = 5
#~~< Available_DateTime >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Available_DateTime = New-Object System.Windows.Forms.DateTimePicker
$Available_DateTime.CustomFormat = "yyyy-MM-dd hh:mm tt"
$Available_DateTime.Format = [System.Windows.Forms.DateTimePickerFormat]::Custom
$Available_DateTime.Location = New-Object System.Drawing.Point(262, 339)
$Available_DateTime.Size = New-Object System.Drawing.Size(187, 20)
$Available_DateTime.TabIndex = 4
#~~< Label2 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label2 = New-Object System.Windows.Forms.Label
$Label2.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$Label2.Font = New-Object System.Drawing.Font("Tahoma", 10.0, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$Label2.Location = New-Object System.Drawing.Point(62, 392)
$Label2.Size = New-Object System.Drawing.Size(176, 23)
$Label2.TabIndex = 10
$Label2.Text = "Required_Schedule"
#~~< Label4 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label4 = New-Object System.Windows.Forms.Label
$Label4.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$Label4.Font = New-Object System.Drawing.Font("Tahoma", 10.0, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$Label4.Location = New-Object System.Drawing.Point(62, 339)
$Label4.Size = New-Object System.Drawing.Size(176, 23)
$Label4.Text = "Available_Schedule"
#~~< ColID >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$ColID = New-Object System.Windows.Forms.TextBox
$ColID.Location = New-Object System.Drawing.Point(262, 184)
$ColID.Size = New-Object System.Drawing.Size(290, 20)
$ColID.TabIndex = 2
$ColID.Text = ""
#~~< Collection_ID >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Collection_ID = New-Object System.Windows.Forms.Label
$Collection_ID.AccessibleDescription = "Collection ID"
$Collection_ID.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$Collection_ID.Font = New-Object System.Drawing.Font("Tahoma", 10.0, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$Collection_ID.Location = New-Object System.Drawing.Point(62, 184)
$Collection_ID.Size = New-Object System.Drawing.Size(176, 21)
$Collection_ID.Text = "SCCM Collection ID"
#~~< appname >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$appname = New-Object System.Windows.Forms.TextBox
$appname.Location = New-Object System.Drawing.Point(262, 129)
$appname.Size = New-Object System.Drawing.Size(290, 20)
$appname.TabIndex = 1
$appname.Text = ""
$appname.BackColor = [System.Drawing.SystemColors]::Window
#~~< Application_Name >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Application_Name = New-Object System.Windows.Forms.Label
$Application_Name.AccessibleDescription = "SCCM Application Displayed on the console"
$Application_Name.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$Application_Name.Font = New-Object System.Drawing.Font("Tahoma", 10.0, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$Application_Name.Location = New-Object System.Drawing.Point(62, 129)
$Application_Name.Size = New-Object System.Drawing.Size(176, 21)
$Application_Name.Text = "SCCM Application Name"
#~~< intent >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$intent = New-Object System.Windows.Forms.ComboBox
$intent.FormattingEnabled = $true
$intent.Location = New-Object System.Drawing.Point(262, 290)
$intent.Size = New-Object System.Drawing.Size(187, 21)
$intent.TabIndex = 3
$intent.Text = ""
#$intent.Items.AddRange([System.Object[]](@("Available", "Required")))
$intent.SelectedIndex = -1
#~~< IntentGo >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$IntentGo = New-Object System.Windows.Forms.Button
$IntentGo.Location = New-Object System.Drawing.Point(477, 290)
$IntentGo.Size = New-Object System.Drawing.Size(75, 25)
$IntentGo.TabIndex = 21
$IntentGo.Text = "Go...!"
$IntentGo.UseVisualStyleBackColor = $true
$IntentGo.add_Click({IntentGoClick($IntentGo)})
#~~< DeploymentAction >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DeploymentAction = New-Object System.Windows.Forms.ComboBox
$DeploymentAction.FormattingEnabled = $true
$DeploymentAction.Location = New-Object System.Drawing.Point(262, 235)
$DeploymentAction.Size = New-Object System.Drawing.Size(187, 21)
$DeploymentAction.TabIndex = 3
$DeploymentAction.Text = ""
$DeploymentAction.Items.AddRange([System.Object[]](@("Install", "Uninstall")))
$DeploymentAction.SelectedIndex = -1
#~~< DeploymentActionLBL >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DeploymentActionLBL = New-Object System.Windows.Forms.Label
$DeploymentActionLBL.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$DeploymentActionLBL.Font = New-Object System.Drawing.Font("Tahoma", 10.0, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$DeploymentActionLBL.Location = New-Object System.Drawing.Point(62, 235)
$DeploymentActionLBL.Size = New-Object System.Drawing.Size(176, 23)
$DeploymentActionLBL.Text = "Deployment Action"
#~~< DeploymentactionGo >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$DeploymentactionGo = New-Object System.Windows.Forms.Button
$DeploymentactionGo.Location = New-Object System.Drawing.Point(477, 235)
$DeploymentactionGo.Size = New-Object System.Drawing.Size(75, 25)
$DeploymentactionGo.TabIndex = 21
$DeploymentactionGo.Text = "Go...!"
$DeploymentactionGo.UseVisualStyleBackColor = $true
$DeploymentactionGo.add_Click({DeploymentactionGoClick($DeploymentactionGo)})
#~~< Label3 >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Label3 = New-Object System.Windows.Forms.Label
$Label3.AccessibleDescription = "Available or Required"
$Label3.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$Label3.Font = New-Object System.Drawing.Font("Tahoma", 10.0, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$Label3.Location = New-Object System.Drawing.Point(62, 290)
$Label3.Size = New-Object System.Drawing.Size(176, 23)
$Label3.Text = "Deployment_Intent"
#~~< OutMTWReboot >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$OutMTWReboot = New-Object System.Windows.Forms.CheckBox
$OutMTWReboot.Location = New-Object System.Drawing.Point(262, 503)
$OutMTWReboot.Size = New-Object System.Drawing.Size(139, 24)
$OutMTWReboot.TabIndex = 22
$OutMTWReboot.Text = "Reboot....!"
$OutMTWReboot.UseVisualStyleBackColor = $true
#~~< outMTWSWinstall >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$outMTWSWinstall = New-Object System.Windows.Forms.CheckBox
$outMTWSWinstall.Location = New-Object System.Drawing.Point(262, 450)
$outMTWSWinstall.Size = New-Object System.Drawing.Size(139, 24)
$outMTWSWinstall.TabIndex = 12
$outMTWSWinstall.Text = "Software Installation"
$outMTWSWinstall.UseVisualStyleBackColor = $true
#~~< Outsidemtw >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Outsidemtw = New-Object System.Windows.Forms.Label
$Outsidemtw.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$Outsidemtw.Font = New-Object System.Drawing.Font("Tahoma", 10.0, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Point, ([System.Byte](0)))
$Outsidemtw.Location = New-Object System.Drawing.Point(62, 470)
$Outsidemtw.Size = New-Object System.Drawing.Size(176, 24)
$Outsidemtw.TabIndex = 17
$Outsidemtw.Text = "Outside Maintence Window"
#~~< whatif >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$whatif = New-Object System.Windows.Forms.Button
$whatif.Location = New-Object System.Drawing.Point(622, 565)
$whatif.Size = New-Object System.Drawing.Size(75, 23)
$whatif.TabIndex = 0
$whatif.Text = "Validate"
$whatif.UseVisualStyleBackColor = $true
$whatif.add_Click({WhatifClick($whatif)})
#~~< deploy >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$deploy = New-Object System.Windows.Forms.Button
$deploy.Location = New-Object System.Drawing.Point(776, 565)
$deploy.Size = New-Object System.Drawing.Size(75, 23)
$deploy.TabIndex = 8
$deploy.Text = "Deploy..!"
$deploy.UseVisualStyleBackColor = $true
$deploy.add_Click({DeployClick($deploy)})
#~~< Exit >~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
$Exit = New-Object System.Windows.Forms.Button
$Exit.Location = New-Object System.Drawing.Point(916, 565)
$Exit.Size = New-Object System.Drawing.Size(75, 23)
$Exit.TabIndex = 9
$Exit.Text = "Exit"
$Exit.UseVisualStyleBackColor = $true
$Exit.add_Click({ExitClick($Exit)})
#--------------------------------------------------------------------------------------------------

$Form1.Controls.Add($Approverlist)
$Form1.Controls.Add($Approver)
$Form1.Controls.Add($Log)
$Form1.Controls.Add($Exit)
$Form1.Controls.Add($ColID)
$Form1.Controls.Add($appname)
$Form1.Controls.Add($deploy)
$Form1.Controls.Add($Collection_ID)
$Form1.Controls.Add($Application_Name)
$Form1.Controls.Add($whatif)
$Form1.Controls.Add($CRQ_REQ)
$Form1.Controls.Add($CRQ_REQ_NO)
$Form1.Controls.Add($DeploymentAction)
$Form1.Controls.Add($DeploymentActionLBL)
$Form1.Controls.Add($DeploymentactionGo)
#endregion

#region Custom Code
                # Site configuration
                $SiteCode = "ABC" # Site code 
                $ProviderMachineName = "Server-25" # SMS Provider machine name

                # Customizations
                $initParams = @{}
                #$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
                #$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

                # Do not change anything below this line

                # Import the ConfigurationManager.psd1 module 
                if((Get-Module ConfigurationManager) -eq $null) {
                    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
                }

                # Connect to the site's drive if it is not already present
                if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
                    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
                }

                # Set the current location to be the site code.
                Set-Location "$($SiteCode):\" @initParams

      # email function below 

Function Send-Mail
{
    Param (
        [Parameter(Mandatory=$False, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true, 
                   Position=0,
                   ParameterSetName='Parameter Set 1')]
        [ValidateNotNullOrEmpty()] 
        [string[]]$To,

        [Parameter(Mandatory=$False,
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
                   [Alias("Copy")]
        [string[]]$CC,
        [Parameter(Mandatory=$False,
                   Position=1,
                   ParameterSetName='Parameter Set 1')]
                   [Alias("Black")]
        [string[]]$BCC,
        [Parameter(Mandatory=$false,
                   Position=2,
                   ParameterSetName='Parameter Set 1')]
        $From = "$($env:username)@domain.com" ,
        $subject = "Mail triggered from Script",
        $body,
        $Attachment,
        [Alias("HTML")]
        [Switch]$AsHtml,

        $SMTPServer = "App-mail.abc.corp.bce.ca"
    )

    Begin
    {
        $MailMessage = New-Object System.Net.Mail.MailMessage 
		$SMTPClient = New-Object System.Net.Mail.smtpClient

        $MailMessage.Sender = $From
		$MailMessage.From = $From

        Function Check-Format
        {
            Param ($Data)
            
            If ($Data -contains ",")
            {
                $Data = $Data.Split(",")
            }
            ElseIf ($Data -contains ";")
            {
                $Data = $Data.Split(";")
            }

            Return $Data
        }
    }

    Process
    {

        $SMTPClient.host = $SMTPServer
		        
        If ($subject)
        {
		    $MailMessage.Subject = $subject
        }

		If ($Attachment)
        {
            Foreach ($File in $Attachment)
            {
                $MailMessage.Attachments.Add($File)
            }
        }

        If ($To)
        {
            $To = Check-Format -Data $To
		    $To | % {$MailMessage.To.add($_)}
        }	
        If ($CC)
        {
            $CC = Check-Format -Data $CC
            $CC | % {$MailMessage.CC.add($_)}
        }
        If ($Bcc)
        {
            $BCC = Check-Format -Data $BCC
		    $BCC | % {  $MailMessage.BCC.add($_)}
        }
		
        If ($Body)
        {
            $MailMessage.Body = $body
		}
        
        If ($AsHtml)
        {
            $MailMessage.IsBodyHtml = $true
		}
    }

    End
    {
        $SMTPClient.Send($MailMessage)
    }
}

#endregion

#region Event Loop

function Main{
[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::Run($Form1)
}

#endregion

#endregion

#region Event Handlers

function DeploymentactionGoClick ($object){

If($DeploymentAction.SelectedItem -eq "uninstall")
    {
    $intent.Items.AddRange([System.Object[]](@("Required"))) 
    $Form1.Controls.Add($intent)
    $Form1.Controls.Add($IntentGo)
    $Form1.Controls.Add($Label3)

    


    }
    elseif($DeploymentAction.SelectedItem -eq "Install")
    {
    $intent.Items.AddRange([System.Object[]](@("Available","Required"))) 
    $Form1.Controls.Add($intent)
    $Form1.Controls.Add($IntentGo)
    $Form1.Controls.Add($Label3)
    }

}



function IntentGoClick( $object ){

If ($intent.SelectedItem -eq "Available" )
    {
    
    $Form1.Controls.Add($Available_DateTime)
    $Form1.Controls.Add($AvailableUTC)
    $Form1.Controls.Add($Label4) # Available schedule Lable

    }
    elseif($intent.SelectedItem -eq "Required")
    {
    $Form1.Controls.Add($Available_DateTime)
    $Form1.Controls.Add($AvailableUTC)
    $Form1.Controls.Add($Label4) # Available schedule Lable

        $Form1.Controls.Add($RequiredUTC)
        $Form1.Controls.Add($Required_DateTime)
        $Form1.Controls.Add($outMTWSWinstall)
        $Form1.Controls.Add($Outsidemtw)
        $Form1.Controls.Add($OutMTWReboot)
        $Form1.Controls.Add($Label2) #Required schedule label

}
}
function WhatifClick( $object ){
If(!($Approverlist.SelectedItem -eq $null )){
$global:Flag = 'No'


            $CM_App = Get-CMApplication -Name $appname.Text 
            $CM_Col = Get-CMDeviceCollection -Id $ColID.Text

 If($CM_App){
         If($CM_Col){
           If($DeploymentAction.SelectedItem -eq "Install" -or $DeploymentAction.SelectedItem -eq "Uninstall"){
                     If ($intent.SelectedItem -eq "Available" -or $intent.SelectedItem -eq "Required"){                              
                               $global:Flag = 'yes'
                               If($intent.SelectedItem -eq "Available"){
                                $Log.Text = "This section is displaying Info based on the given input" 
                                $log.AppendText("`n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                $log.AppendText("`nApplication Name = $($CM_App.LocalizedDisplayName)")
                                $log.AppendText("`nApplication Version = $($CM_App.LocalizedDescription)")
                                $log.AppendText("`nCollection Name = $($CM_Col.Name)")
                                $log.AppendText("`nLimiting Collection = $($CM_Col.LimitToCollectionName)")
                                $log.AppendText("`nDeploymentAction =$($DeploymentAction.SelectedItem)")
                                $log.AppendText("`nDeployment =$($intent.SelectedItem)")
                                $log.AppendText("`nAvailable Time $(($Available_DateTime.Value).AddMinutes(30))") 
                                $log.AppendText("`n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                $log.AppendText("`nDeployment Will be created with 30 minutes Grace period, once email is received, approver has to approve it in the email.")
                                $log.AppendText("`n------------------------------------------------------")                              
                               } 
                               elseif($intent.SelectedItem -eq "Required"){
                                
                                $Log.Text = "This section is displaying Info based on the given input" 
                                $log.AppendText("`n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                $log.AppendText("`nApplication Name = $($CM_App.LocalizedDisplayName)")
                                $log.AppendText("`nApplication Version = $($CM_App.LocalizedDescription)")
                                $log.AppendText("`nCollection Name = $($CM_Col.Name)")
                                $log.AppendText("`nLimiting Collection = $($CM_Col.LimitToCollectionName)")
                                $log.AppendText("`nDeployment =$($intent.SelectedItem)")
                                $log.AppendText("`nDeploymentAction =$($DeploymentAction.SelectedItem)")
                                $log.AppendText("`nAvailable Time $(($Available_DateTime.Value).AddMinutes(30))")
                                $log.AppendText("`nRequired Time $(($Required_DateTime.value).AddMinutes(30))")
                                $log.AppendText("`n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
                                $log.AppendText("`nDeployment Will be created with 30 minutes Grace period, once email is received, approver has to approve it in the email.")
                                $log.AppendText("`n------------------------------------------------------")
                                }
                                   }else{$log.Clear();$log.AppendText("`nSelect Intent")  }
                                }else{$log.Clear();$log.AppendText("`nSelect Deployment Action") }
                             }else {$log.Clear();$log.AppendText("`nEnter Valid Collection ID") }
         }else {$log.Clear();$log.AppendText("`nEnter Valid Application Name") }                       
}else{$log.Clear(); $log.AppendText("`nSelect Approver email ID from the drop down...!")}
}

function ExitClick( $object ){
$Form1.Close()
}

function DeployClick( $object ){
        Try{

      If ($Flag -eq 'Yes'){
                
$CM_App = Get-CMApplication -Name $appname.Text
$CM_Col = Get-CMDeviceCollection -Id $ColID.Text


If ($intent.SelectedItem -eq "Available" )
    {
    $log.AppendText("`nCreating Available Deployment with Below details")
    
    If($AvailableUTC.CheckState -eq "UnChecked"){$cm_Utc = 'LocalTime'}elseif($AvailableUTC.CheckState -eq "Checked"){$cm_Utc = 'Utc'}
$CM_GUID =$(New-Guid)
$CM_time = $(Get-Date)
    $CM_comment = @"
CRQ_WO_REQ = $($CRQ_REQ.Text)
Tool_Guid = $CM_GUID
Creation_Time = $CM_time
User_ID = $($env:USERNAME)
"@ 
    $CM_newDeploy = New-CMApplicationDeployment -Name $CM_App.LocalizedDisplayName -DeployAction $($DeploymentAction.SelectedItem) -DeployPurpose Available -CollectionId $CM_Col.CollectionID -UserNotification DisplaySoftwareCenterOnly -TimeBaseOn $cm_Utc -AvailableDateTime $($Available_DateTime.Value).AddMinutes(30) -PersistOnWriteFilterDevice $False -Comment $CM_comment
   
    $log.AppendText("`n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    $log.AppendText("`nApplication Name = $($cm_newDeploy.ApplicationName)")
    $log.AppendText("`nApplication Version = $($CM_App.LocalizedDescription)")
    $log.AppendText("`nCollection Name = $($cm_newDeploy.CollectionName)")
    $log.AppendText("`nLimiting Collection = $($CM_Col.LimitToCollectionName)")
    $log.AppendText("`nDeployment =$($intent.SelectedItem)")
    $log.AppendText("`nDeploymentAction =$($DeploymentAction.SelectedItem)")
    $log.AppendText("`nAvailable Time $($cm_newDeploy.StartTime)")
    $log.AppendText("`n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    
    }
    elseif($intent.SelectedItem -eq "Required")
    {
    $log.AppendText("`nCreating Required/Push Deployment with Below details")

    If($RequiredUTC.CheckState -eq "UnChecked"){$CM_Utc = 'LocalTime'}elseif($RequiredUTC.CheckState -eq "Checked"){$cm_Utc = 'Utc'}
    If($outMTWSWinstall.CheckState -eq "Unchecked"){$CM_mtwoverride = $False }elseif($outMTWSWinstall.CheckState -eq "checked"){$CM_mtwoverride = $true}
    If($OutMTWReboot.CheckState -eq "Unchecked"){$CM_OutMTWReboot = $False }elseif($OutMTWReboot.CheckState -eq "checked"){$CM_OutMTWReboot = $true}

$CM_GUID =$(New-Guid)
$CM_time = $(Get-Date)
    $CM_comment = @"
CRQ_WO_REQ = $($CRQ_REQ.Text)
Tool_Guid = $CM_GUID
Creation_Time = $CM_time
User_ID = $($env:USERNAME)
"@  
    $cm_newDeploy = New-CMApplicationDeployment -Name $CM_App.LocalizedDisplayName -DeployAction $($DeploymentAction.SelectedItem) -DeployPurpose Required -CollectionId $CM_Col.CollectionID -UserNotification HideAll -TimeBaseOn $cm_Utc -AvailableDateTime $($Available_DateTime.Value).AddMinutes(30) -DeadlineDateTime $($Required_DateTime.Value).AddMinutes(30) -PersistOnWriteFilterDevice $False -OverrideServiceWindow $CM_mtwoverride -Comment $CM_comment -RebootOutsideServiceWindow $cm_OutMTWReboot

    $log.AppendText("`n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    $log.AppendText("`nApplication Name = $($cm_newDeploy.ApplicationName)")
    $log.AppendText("`nApplication Version = $($CM_App.LocalizedDescription)")
    $log.AppendText("`nCollection Name = $($cm_newDeploy.CollectionName)")
    $log.AppendText("`nLimiting Collection = $($CM_Col.LimitToCollectionName)")
    $log.AppendText("`nDeployment = $($intent.SelectedItem)")
    $log.AppendText("`nDeploymentAction =$($DeploymentAction.SelectedItem)")
    $log.AppendText("`nAvailable Time $($cm_newDeploy.StartTime)")
    $log.AppendText("`nRequired Time $($cm_newDeploy.EnforcementDeadline)")
    $log.AppendText("`n~~~~~~~~~~~~~~~~~~~~~~~~~~~~~")
    
    }
    $obj = @" 
<h4>Hi Team,</h4>
<h4>Deployment is created with 30 minutes Grace period, Approver has to approve it via email within next 30 minutes</h4>
<h3> SCCM Application Deployment summary</h3>

<table Border="1">
	<tr>
		<th>CRQ or REQ or WO</th>
		<td>$($CRQ_REQ.Text)</td>
	</tr>
	<tr>
		<th>Application Name</th>
		<td>$($cm_newDeploy.ApplicationName)</td>
	</tr>
    <tr>
		<th>Collection name</th>
		<td>$($cm_newDeploy.CollectionName)</td>    
	</tr>
    <tr>
		<th>Collection ID</th>
		<td>$($cm_newDeploy.TargetCollectionID)</td>    
    </tr>
     <tr>
		<th>Collection Member Count</th>
		<td>$($(Get-CMDeviceCollection -Id $($cm_newDeploy.TargetCollectionID)).MemberCount)</td>    
    </tr>
     <tr>
		<th>Deploy Action</th>
		<td>$($DeploymentAction.SelectedItem)</td>    
    </tr>
    <tr>
		<th>Deploy Intent</th>
		<td>$($intent.SelectedItem)</td>    
    </tr>
     <tr>
		<th>Available Time</th>
		<td>$($cm_newDeploy.StartTime)</td>    
    </tr>
 
     <tr>
		<th>Required Time</th>
		<td>$($cm_newDeploy.EnforcementDeadline)</td>       
    </tr>
    
    <tr>
		<th>App Deployment Id</th>
		<td>$($cm_newDeploy.AssignmentID)</td>    
	</tr>

</table>
<p></p>
<p></p>
<p></p>

<h4>Tool Data </h4>
<p></p>
<p></p>
<p></p>

<table border="2" >
<tr>
	<th>Tool GUID</th>
	<th>Time</th>
	<th>UserID</th>
</tr>
<tr>
	<td>$CM_GUID</td>
	<td>$CM_time</td>
	<td>$($env:USERNAME)</td>
</tr>
</table>

<p>Thanks,</p>
<p>$($env:USERNAME)</p>
"@
                            
Send-Mail -To $($Approverlist.SelectedItem) -CC "SCCM.Ops.copr@abc.com" -subject "Application Deployment Summary - $($cm_newDeploy.ApplicationName)" -body $obj -From "MECM_Automation@abc.com" -AsHtml


}else { $log.clear(); $log.AppendText("`nValidate Button should be Satisfied to Deploy.......!"); $log.AppendText("`nOr Please Click the Exit button to Exit the App") }

}catch{ $log.Clear(); $log.AppendText( $($_.exception.message))
    }



}





Main # This call must remain below all other event functions

#endregion
# ========================================================
#
# Script Information
#
# Title: SCCM APplication Deployment Toll for validation 
# Author: PERSONAL\master
# Originally created: 12-09-2023 - 01:37:19
# Original path: C:\Users\master\Documents\SccmApplicationDeploymentTool.ps1
# Description: 
# 
# ========================================================
