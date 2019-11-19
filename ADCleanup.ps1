<#-------------------------
Name: ADCleanup_UI_0_5.ps1
Author: Cameron Wisniewski
Date: 1/15/19
Comment: Utility used for cleaning up old computers in Active Directory
-------------------------#>

#Requires –Modules ActiveDirectory  

$Version = "v0.5"

#Load .NET Elements for the UI
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")

#Load Required Powershell Modules
Import-Module ActiveDirectory

#Set some variables we'll need a few times later or that we might want to change at some point:
#--------------------------
$CompanyName = "<COMPANYNAME>"
$LogonCutoff = (Get-Date).AddDays(-90) #Edit the number, default -90, to change the LastLogonDate cutoff
$TargetOU = "<SEARCHBASE>" #Edit this to change the searchbase. Generally, this shouldn't be changed. 
$CredentialsAccessGroup = "<RUNACCESSGROUP>" #Defines which AD group a user must be a member of to make changes to AD, allows you to run the tool.
$OptionsAccessGroup  = "<FULLACCESSGROUP>" #Defines which AD group a user must be a member of to access the options pane to change the Description filter/SearchBase/LastLogonDate
$LogFolderPath = "<LOGFOLDER>" #Location of the logs folder
$LogoPath = "<COMPANYLOGOPATH>"
$Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($PSHOME + "\powershell.exe") #Sets the icon for the app
$ErrorType = "Ok" #Set the button default for the error popup. Generally, this shouldn't be changed.

#Set some stuff to manipulate date/time
$DateTime = (Get-Date).ToString()
$Date = (Get-Date -UFormat "%Y%m%d").ToString()
$Year = (Get-Date -UFormat %Y).ToString()
$Month = (Get-Date -Format MMMM).ToString()
$Time = (Get-Date -UFormat "%H%M%S")
$ErrorActionPreference = "Continue" 
$LogFilePath = "$LogFolderPath\ADCleanup.$Date.$Time.log" #Set the name of the log file based on the date/time

#Start logging the output of the script
New-Item $LogFilePath -ItemType File
Start-Transcript -Path $LogFilePath -Append

#Credentials Button - Gather alternate credentials to ensure these actions are run with the proper permissions
$CredentialsButton_Click = {
    #Gather credentials with the Get-Credential cmdlet to keep things secure
    $global:AltCredentials = Get-Credential -Message "Enter credentials with AD write access."
    $global:AltUN = $AltCredentials.UserName
    Write-Host "LOG: Checking credentials for $AltUN."

    #See if those credentials are part of the security group that allows AD access
    #The below code is currently buggy for some reason? - CJW 9-20-19
    #$CredsCheck = Get-ADPrincipalGroupMembership -Identity $AltUN -Credential $AltCredentials | Where -Property Name -EQ $CredentialsAccessGroup
    $CredsCheck = (Get-ADUser -Identity $AltUN -Properties *).MemberOf | ForEach-Object {
        if($_.Contains($CredentialsAccessGroup)){$True}
    }

    #Throw and error and clear relevant variables if the credentials don't have the right permissions.
    #Clearing out the variables stops you from getting locked out by furious clicking and also allows
    #for these variables to be used as checks later.
    if ($CredsCheck) {
        Write-Host "LOG: Credentials for user $AltUN have the necessary privileges and will be used."
        } else {
        Write-Host "LOG: The specified credentials do not have the necessary privileges."
        Clear-Variable -Name "AltCredentials" -Scope Global
        Clear-Variable -Name "AltUN" -Scope Global
        Call-Error -ErrorTitle "Credentials Error" -ErrorString "The specified credentials do not have the necessary privileges."
        return
        }        
    $StatusBar.Text = "Credentials for user $AltUN will be used."
    $UnderGridInstructions.Text = "Click the Report button to generate the list of applicable computers."
    }

#Report Button - Generate a report containing computers with LastLogonDate exceeding the X day cutoff specified in the LogonCutoff variable. 
$ReportButton_Click = {
    if(!$AltCredentials) {
        Call-Error -ErrorTitle "Credentials Error" -ErrorString "Please enter your credentials using the Credentials button."
        return
        }
    $ADCleanup.Controls.Remove($UnderGridInstructions)
    $ADCleanup.Controls.Remove($ProgressBar)
    $ADCleanup.Controls.Remove($ProgressBarLabel)
    $ADCleanup.Controls.Add($DataGrid)
    $ADCleanup.Controls.Add($ReportExportButton)
    $DataGrid.Rows.Clear()
    $global:OldComputers = Get-ADComputer -SearchBase $TargetOU -Properties LastLogonDate, Description, CanonicalName -Filter {(LastLogonDate -lt $LogonCutoff)} -Credential $AltCredentials | Where Description -Like "$TargetDescription"
    Write-Host "LOG: Generating report from SearchBase $TargetOU with a LastLogonDate less than $LogonCutoff with description like: $TargetDescription"
    $OldComputers | ForEach-Object {
        Write-Host "$_.Name,$_.LastLogonDate,$_.Description,$_.CanonicalName"
        $DataGrid.Rows.Add($_.Name,$_.LastLogonDate,$_.Description,$_.CanonicalName) 
        }
    $StatusBar.Text = "Computers with LastLogonDate earlier than $LogonCutoff are now displayed."
    }

#Process Button - Process each computer object by disabling it, changing it's description to it's CN, and moving it.
#                 The requisite OU is also created if it hasn't been. This requires for the first two steps to have
#                 been run in order to work properly
$ProcessButton_Click = {
    #Check that credentials have been entered
    if(!$AltCredentials) {
        Call-Error -ErrorTitle "Credentials Error" -ErrorString "Please enter your credentials using the Credentials button."
        return
        }

    #Check that a report has actually been run. 
    if(!$OldComputers){
        Call-Error -ErrorTitle "Report Error" -ErrorString "You cannot process results without first generating a list of computers."
        return
        }

    #Verify that you actually want to proceed with processing the computers
    Call-Error -ErrorTitle "Proceed?" -ErrorString "These computers will be disabled and moved in AD. Would you like to proceed?" -ErrorType "OkCancel"

    #Calling this "Advanced Return" bit to make the cancel button actually work. I'm not sure there's a better way to do this, since calling return as a result of
    #cancel in the switch statement just returns out of the Call-Error function.
    if($AdvReturn) {
        Clear-Variable AdvReturn -Scope "Global"
        return
        }
    Write-Host "LOG: Checking for necessary OU structure."

    #Check if the OU for the year's computer cleanup exists already.
    $YearOUExist = Get-ADOrganizationalUnit -Identity "OU=$Year,OU=Computer Deletion,DC=us,DC=crowncastle,DC=com"
    
    #If it doesn't exist, create it according to the date/time variables specified at the beginning of the script
    if(!$YearOUExist){
        Write-Host "LOG: Attempting to create necessary OU structure for the year."
        New-ADOrganizationalUnit -Name $Year -Path "OU=Computer Deletion,DC=us,DC=crowncastle,DC=com" -Credential $AltCredentials -ErrorAction Continue | Write-Host
        } else {
        Write-Host "LOG: OU=$Year,OU=Computer Deletion,DC=us,DC=crowncastle,DC=com already exists."
        }

    #Check if the OU for the month's's computer cleanup exists already.
    $MonthOUExist = Get-ADOrganizationalUnit -Identity "OU=$Month,OU=$Year,OU=Computer Deletion,DC=us,DC=crowncastle,DC=com"

    #If it doesn't exist, create it according to the date/time variables specified at the beginning of the script
    if(!$MonthOUExist){
        Write-Host "LOG: Attempting to create necessary OU structure for the month."
        New-ADOrganizationalUnit -Name $Month -Path "OU=$Year,OU=Computer Deletion,DC=us,DC=crowncastle,DC=com" -Credential $AltCredentials -ErrorAction Continue
        } else {
        Write-Host "LOG: OU=$Month,OU=$Year,OU=Computer Deletion,DC=us,DC=crowncastle,DC=com already exists."
        }
    Write-Host "LOG: Attempting operations on computers in the report."

    #Remove data grid
    $ADCleanup.Controls.Remove($DataGrid)
    $ADCleanup.Controls.Remove($ReportExportButton)

    #Add progress bar & necessary vars
    $ADCleanup.Controls.Add($ProgressBar)
    $ADCleanup.Controls.Add($ProgressBarLabel)
    $SelectedCount = $DataGrid.SelectedRows.Count
    $CurrentCount = 0
    [int]$ProcessPercentage = 0

    #Process the computer objects
    $DataGrid.SelectedRows | ForEach-Object {
        $Computer = $OldComputers[$_.Index].Name
        $Description = $OldComputers[$_.Index].CanonicalName
        $GUID = $OldComputers[$_.Index].ObjectGUID

        #Set the description to the CN and disable it
        Write-Host "LOG: Attempting Disable and Set Descriptions actions on $Computer."
        Set-ADComputer $Computer -Description $Description -Enabled $false -Credential $AltCredentials

        #Move it to the proper OU
        Write-Host "LOG: Move action attempted on $Computer to: OU=$Month,OU=$Year,OU=Computer Deletion,DC=us,DC=crowncastle,DC=com"
        Move-ADObject $GUID -TargetPath "OU=$Month,OU=$Year,OU=Computer Deletion,DC=us,DC=crowncastle,DC=com" -Credential $AltCredentials

        #Increment Counter
        $CurrentCount += 1
        $ProcessPercentage = ($CurrentCount/$SelectedCount)*100
        $ProgressBar.Value = $ProcessPercentage
        $StatusBar.Text = "Processing $ProcessPercentage% complete. ($CurrentCount/$SelectedCount)"
        $ProgressBarLabel.Text = "Processing $ProcessPercentage% complete. ($CurrentCount/$SelectedCount)"
        $ADCleanup.Refresh()
        }
    Write-Host "LOG: Operations attempt complete."

    #Call an error pop to notify that the operation is complete, then close the app.
    Call-Error -ErrorTitle "Operation complete." -ErrorString "These computers have been Disabled, have had their Descriptions set to their previous AD location, and have been Moved to the Computer Deletion OU. Please verify the results."
    Clear-Variable OldComputers
    }


#Report Export Button - Exports the list of computers to be processed to a .csv file in a location of the user's choice. 
$ReportExportButton_Click = {
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    [void]$FolderBrowser.ShowDialog()
    $ExportPath = $FolderBrowser.SelectedPath
    $OldComputers | Export-Csv "$ExportPath\ADCleanup_Report_$Date.csv"
    Call-Error -ErrorTitle "Report created!" -ErrorString "A .CSV file containing this list has been exported to $ExportPath\ADCleanup_Report_$Date.csv."
    }

#Options Button - Displays the options pane.
$OptionsButton_Click = {
    if($AltCredentials) {
        if((Get-ADPrincipalGroupMembership -Identity $AltUN -Credential $AltCredentials | Where -Property Name -EQ $OptionsAccessGroup)){
            [void] $OptionsPane.ShowDialog()
        } else {
            Call-Error -ErrorTitle "Insufficient Credentials" -ErrorString "Credentials for user $AltUN do not have the necessary privileges to use this function."
        }
    } else {
        Call-Error -ErrorTitle "Credentials Error" -ErrorString "Please enter your credentials using the Credentials button."
    }
}

$OptionsApplyButton_Click = {
    #Set variables
    $script:TargetDescription = $OptionsPaneDescriptionField.Text
    $script:TargetOU = $OptionsPaneSearchBaseField.Text
    $script:LogonCutoff = $OptionsPaneLogonDateCutoffField.Text
    #Close options pane
    $OptionsPane.Close()
}

#Help Button - Displays the KB article for this tool if the user has questions about it's functionality.
#              Not really functional since the article doesn't exist yet. 
$HelpButton_Click = {
    Start-Process -FilePath "http://intranet.crowncastle.com/private/it/enterpriseops/hd/desktop/adcleanup_wiki/AD%20Cleanup%20Procedures.aspx"
    $StatusBar.Text = "Documentation opened in browser."
    }

#Exit Button - Closes the application
$ExitButton_Click = {
    $ADCleanup.Close()
    }

#Error Window - Wrote this up as a function to call in a number of different cases for efficiency
function Call-Error() {
    Param(
    [Parameter(Mandatory=$True)]
    [string]$ErrorTitle,
    [Parameter(Mandatory=$True)]
    [string]$ErrorString,
    [ValidateSet("OK","OkCancel")]
    [string]$ErrorType = "OK"
    )
    $ErrorPopup = [System.Windows.Forms.MessageBox]::Show("$ErrorString","$ErrorTitle","$ErrorType","Warning")
    switch($ErrorPopup) {
        OK {}
        Cancel {$global:AdvReturn = 1}
        }    
    }

#region Primary UI Elements - Draws the primary UI. Stuff is named accordingly and is generally in the order you see it in the app.
Write-Host "LOG: Drawing UI"
$ADCleanup = New-Object System.Windows.Forms.Form
$ADCleanup.Text = "$CompanyName - AD Computer Cleanup Utility $Version"
$ADCleanup.Size = New-Object System.Drawing.Size(1000,700)
$ADCleanup.StartPosition = "CenterScreen"
$ADCleanup.FormBorderStyle = "Fixed3D"
$ADCleanup.MaximizeBox = $False
$ADCleanup.Icon = $Icon

$CCILogo = [System.Drawing.Image]::FromFile($LogoPath)
$CCILogo_Embed = New-Object Windows.Forms.PictureBox
$CCILogo_Embed.Width = $CCILogo.Size.Width
$CCILogo_Embed.Height = $CCILogo.Size.Height
$CCILogo_Embed.Size = New-Object System.Drawing.Size(240,64)
$CCILogo_Embed.Location = New-Object Drawing.Point((($ADCleanup.Size.Width/2)-($CCILogo.Size.Width/2)),20)
$CCILogo_Embed.Image = $CCILogo
$ADCleanup.Controls.Add($CCILogo_Embed)

$Instructions = New-Object System.Windows.Forms.Label
$Instructions.Location = New-Object System.Drawing.Size(5,($CCILogo_Embed.Size.Height+30))
$Instructions.Size = New-Object System.Drawing.Size(($ADCleanup.Size.Width-20),60)
$Instructions.Font = New-Object System.Drawing.Font("Calibri", 10, [System.Drawing.FontStyle]::Regular)
$Instructions.Text =  "This utility can be used to perform routine AD computer maintainence on computers that have not been logged into within the past 90 days and which do not currently have a description set. First, enter the appropriate credentials using the Credentials button. From there, gather a list of these computers with the Report button. After selecting each of the computers that should be processed, click the Disable, set Description, and Move button."
$ADCleanup.Controls.Add($Instructions)

$DataGrid = New-Object System.Windows.Forms.DataGridView
$DataGrid.Location = New-Object System.Drawing.Size(5,($Instructions.Location.Y+60))
$DataGrid.Size = New-Object System.Drawing.Size(($ADCleanup.Size.Width-30),($ADCleanup.Size.Height*0.54))
$DataGrid.ColumnCount = 4
$DataGrid.ColumnHeadersVisible = $true
$DataGrid.ReadOnly = $true
$DataGrid.AllowUserToResizeRows = $False
$DataGrid.AllowUserToAddRows = $false
$DataGrid.SelectionMode = "FullRowSelect"
$DataGrid.Columns[0].Name = "Computer Name"
$DataGrid.Columns[1].Name = "LastLogonDate"
$DataGrid.Columns[2].Name = "Description"
$DataGrid.Columns[3].Name = "AD Location"
$DataGrid.Columns[0].Width = ($DataGrid.Size.Width * 0.1)
$DataGrid.Columns[1].Width = ($DataGrid.Size.Width * 0.3)
$DataGrid.Columns[2].Width = ($DataGrid.Size.Width * 0.1)
$DataGrid.Columns[3].Width = ($DataGrid.Size.Width * 0.43)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Size(40,($ADCleanup.Size.Height/2-50))
$ProgressBar.Size = New-Object System.Drawing.Size(($ADCleanup.Size.Width-100),100)
$ProgressBar.Name = "ProgressBar"
$ProgressBar.Value = 0
$ProgressBar.Style = "Continuous"

$ProgressBarLabel = New-Object System.Windows.Forms.Label
$ProgressBarLabel.Location = New-Object System.Drawing.Size(($ADCleanup.Size.Width/2-125),($ADCleanup.Size.Height/2-80))
$ProgressBarLabel.Size = New-Object System.Drawing.Size(250,20)
$ProgressBarLabel.Font = New-Object System.Drawing.Font("Calibri", 10, [System.Drawing.FontStyle]::Regular)
$ProgressBarLabel.Text =  "Processing selected computers..."
$ProgressBarLabel.AutoSize = $false
$ProgressBarLabel.TextAlign = "MiddleCenter"

$CredentialsButton = New-Object System.Windows.Forms.Button
$CredentialsButton.Location = New-Object System.Drawing.Size(5,($ADCleanup.Size.Height-160))
$CredentialsButton.Size = New-Object System.Drawing.Size(100,100)
$CredentialsButton.Text = "Credentials"
$CredentialsButton.Add_Click($CredentialsButton_Click)
$ADCleanup.Controls.Add($CredentialsButton)

$ReportButton = New-Object System.Windows.Forms.Button
$ReportButton.Location = New-Object System.Drawing.Size(($CredentialsButton.Location.X+105),$CredentialsButton.Location.Y)
$ReportButton.Size = New-Object System.Drawing.Size(100,100)
$ReportButton.Text = "Report"
$ReportButton.Add_Click($ReportButton_Click)
$ADCleanup.Controls.Add($ReportButton)

$ProcessButton = New-Object System.Windows.Forms.Button
$ProcessButton.Location = New-Object System.Drawing.Size(($ReportButton.Location.X+105),$ReportButton.Location.Y)
$ProcessButton.Size = New-Object System.Drawing.Size(100,100)
$ProcessButton.Text = "Disable,`nset Description,`nand Move"
$ProcessButton.Add_Click($ProcessButton_Click)
$ADCleanup.Controls.Add($ProcessButton)

$ReportExportButton = New-Object System.Windows.Forms.Button
$ReportExportButton.Location = New-Object System.Drawing.Size(($ADCleanup.Size.Width-100),($CredentialsButton.Location.Y))
$ReportExportButton.Size = New-Object System.Drawing.Size(75,25)
$ReportExportButton.Text = "Export"
$ReportExportButton.Add_Click($ReportExportButton_Click)

$ExitButton = New-Object System.Windows.Forms.Button
$ExitButton.Location = New-Object System.Drawing.Size(($ADCleanup.Size.Width-100),($ADCleanup.Size.Height-85))
$ExitButton.Size = New-Object System.Drawing.Size(75,25)
$ExitButton.Text = "Exit"
$ExitButton.Add_Click($ExitButton_Click)
$ADCleanup.Controls.Add($ExitButton)

$HelpButton = New-Object System.Windows.Forms.Button
$HelpButton.Location = New-Object System.Drawing.Size(($ExitButton.Location.X-80),$ExitButton.Location.Y)
$HelpButton.Size = New-Object System.Drawing.Size(75,25)
$HelpButton.Text = "Help"
$HelpButton.Add_Click($HelpButton_Click)
$ADCleanup.Controls.Add($HelpButton)

$OptionsButton = New-Object System.Windows.Forms.Button
$OptionsButton.Location = New-Object System.Drawing.Size(($HelpButton.Location.X-80),$HelpButton.Location.Y)
$OptionsButton.Size = New-Object System.Drawing.Size(75,25)
$OptionsButton.Text = "Options"
$OptionsButton.Add_Click($OptionsButton_Click)
$ADCleanup.Controls.Add($OptionsButton)

$OptionsPane = New-Object System.Windows.Forms.Form
$OptionsPane.Text = "Options"
$OptionsPane.Size = New-Object System.Drawing.Size(500,250)
$OptionsPane.StartPosition = "CenterScreen"
$OptionsPane.FormBorderStyle = "Fixed3D"
$OptionsPane.MaximizeBox = $False
$OptionsPane.Icon = $Icon

$OptionsPaneDescriptionLabel = New-Object System.Windows.Forms.Label
$OptionsPaneDescriptionLabel.Location = New-Object System.Drawing.Size(5,10)
$OptionsPaneDescriptionLabel.Size = New-Object System.Drawing.Size(160,20)
$OptionsPaneDescriptionLabel.Font = New-Object System.Drawing.Font("Calibri", 10, [System.Drawing.FontStyle]::Regular)
$OptionsPaneDescriptionLabel.Text =  "Description Filter:"
$OptionsPaneDescriptionLabel.AutoSize = $false
$OptionsPaneDescriptionLabel.TextAlign = "LeftCenter"
$OptionsPane.Controls.Add($OptionsPaneDescriptionLabel)

$OptionsPaneDescriptionField = New-Object System.Windows.Forms.TextBox
$OptionsPaneDescriptionField.Location = New-Object System.Drawing.Size(5,30)
$OptionsPaneDescriptionField.Size = New-Object System.Drawing.Size(470,20)
$OptionsPaneDescriptionField.Font = New-Object System.Drawing.Font("Calibri", 10, [System.Drawing.FontStyle]::Regular)
$OptionsPaneDescriptionField.Text =  $TargetDescription
$OptionsPane.Controls.Add($OptionsPaneDescriptionField)

$OptionsPaneSearchBaseLabel = New-Object System.Windows.Forms.Label
$OptionsPaneSearchBaseLabel.Location = New-Object System.Drawing.Size(5,65)
$OptionsPaneSearchBaseLabel.Size = New-Object System.Drawing.Size(160,20)
$OptionsPaneSearchBaseLabel.Font = New-Object System.Drawing.Font("Calibri", 10, [System.Drawing.FontStyle]::Regular)
$OptionsPaneSearchBaseLabel.Text =  "SearchBase Filter:"
$OptionsPaneSearchBaseLabel.AutoSize = $false
$OptionsPaneSearchBaseLabel.TextAlign = "LeftCenter"
$OptionsPane.Controls.Add($OptionsPaneSearchBaseLabel)

$OptionsPaneSearchBaseField = New-Object System.Windows.Forms.TextBox
$OptionsPaneSearchBaseField.Location = New-Object System.Drawing.Size(5,85)
$OptionsPaneSearchBaseField.Size = New-Object System.Drawing.Size(470,20)
$OptionsPaneSearchBaseField.Font = New-Object System.Drawing.Font("Calibri", 10, [System.Drawing.FontStyle]::Regular)
$OptionsPaneSearchBaseField.Text =  $TargetOU
$OptionsPane.Controls.Add($OptionsPaneSearchBaseField)

$OptionsPaneLogonDateCutoffLabel = New-Object System.Windows.Forms.Label
$OptionsPaneLogonDateCutoffLabel.Location = New-Object System.Drawing.Size(5,125)
$OptionsPaneLogonDateCutoffLabel.Size = New-Object System.Drawing.Size(160,20)
$OptionsPaneLogonDateCutoffLabel.Font = New-Object System.Drawing.Font("Calibri", 10, [System.Drawing.FontStyle]::Regular)
$OptionsPaneLogonDateCutoffLabel.Text =  "Logon Date Cutoff:"
$OptionsPaneLogonDateCutoffLabel.AutoSize = $false
$OptionsPaneLogonDateCutoffLabel.TextAlign = "LeftCenter"
$OptionsPane.Controls.Add($OptionsPaneLogonDateCutoffLabel)

$OptionsPaneLogonDateCutoffField = New-Object System.Windows.Forms.TextBox
$OptionsPaneLogonDateCutoffField.Location = New-Object System.Drawing.Size(5,145)
$OptionsPaneLogonDateCutoffField.Size = New-Object System.Drawing.Size(470,20)
$OptionsPaneLogonDateCutoffField.Font = New-Object System.Drawing.Font("Calibri", 10, [System.Drawing.FontStyle]::Regular)
$OptionsPaneLogonDateCutoffField.Text =  $LogonCutoff
$OptionsPane.Controls.Add($OptionsPaneLogonDateCutoffField)

$OptionsApplyButton = New-Object System.Windows.Forms.Button
$OptionsApplyButton.Location = New-Object System.Drawing.Size(400,175)
$OptionsApplyButton.Size = New-Object System.Drawing.Size(75,25)
$OptionsApplyButton.Text = "Apply"
$OptionsApplyButton.Add_Click($OptionsApplyButton_Click)
$OptionsPane.Controls.Add($OptionsApplyButton)

$UnderGridInstructions = New-Object System.Windows.Forms.Label
$UnderGridInstructions.Size = New-Object System.Drawing.Size(680,100)
$UnderGridInstructions.AutoSize = $False
$UnderGridInstructions.TextAlign = "MiddleCenter"
$UnderGridInstructions.Dock = "Fill"
$UnderGridInstructions.Font = New-Object System.Drawing.Font("Calibri", 24, [System.Drawing.FontStyle]::Regular)
$UnderGridInstructions.Text = "Please enter your credentials using the button below."
$ADCleanup.Controls.Add($UnderGridInstructions)

$StatusBar = New-Object System.Windows.Forms.StatusBar
$StatusBar.Name = "StatusBar"
$StatusBar.Text = "Awaiting input . . ."
$ADCleanup.Controls.Add($StatusBar)
#endregion

#Show UI - Actually draw everything that we just specified above. 
$ADCleanup.TopMost = $false
$ADCleanup.Add_Shown({$ADCleanup.Activate()})
[void] $ADCleanup.ShowDialog()
Stop-Transcript
exit