
$wshell = New-Object -ComObject Wscript.Shell
$wshell.Popup("Please enter in your Domain Admin credentials.  Please remember it should be in the form of DOMAIN\username.",0,"Credentials Needed!",0x0)	
$creds = Get-Credential
 $PSDefaultParameterValues = @{"*-AD*:Credential"=$creds}

#Here we create the connection to the exchange server. Edit with your mailserver info
$ExchangeSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://contosomail.jacontoso.com/PowerShell
Import-PSSession $ExchangeSession

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
[void] [System.Windows.Forms.Appliion]::EnableVisualStyles() 
	
$objForm = New-Object System.Windows.Forms.Form 
$objForm.Text = "contoso Terminated Employee Process Form"
$objForm.Size = New-Object System.Drawing.Size(500,400) 
$objForm.StartPosition = "CenterScreen"
$objForm.MaximizeBox = $False


$objForm.KeyPreview = $True
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
    {$userinput=$UserTextBox.Text;$forwardemail=$ForwardingTextBox.Text;$ticketnumber=$TicketTextBox.Text;$disableuser=$DisableUserCheckbox.Checked;$objForm.Close()}})
$objForm.Add_KeyDown({if ($_.KeyCode -eq "Escape") 
    {$objForm.Close()}})

$Font = New-Object System.Drawing.Font("Verdana",8,[System.Drawing.FontStyle]::Bold) 
#$objForm.Font = $Font 
#VERSION NUMBER
$VersionLabel = New-Object System.Windows.Forms.Label
$VersionLabel.Loion = New-Object System.Drawing.Size(450,10) 
$VersionLabel.Size = New-Object System.Drawing.Size(120,20) 
$VersionLabel.Font = $Font 
$VersionLabel.Text = "V1"
$objForm.Controls.Add($VersionLabel) 

#OK AND CANCEL BUTTONS
$OKButton = New-Object System.Windows.Forms.Button
$OKButton.Loion = New-Object System.Drawing.Size(75,320)
$OKButton.Size = New-Object System.Drawing.Size(75,23)
$OKButton.Text = "OK"
$OKButton.Add_Click({$userinput=$UserTextBox.Text;$ticketnumber=$TicketTextBox.Text;$manageremail=$emailmanagercheckbox.Text;$forwardemail=$ForwardingTextBox.Text;$disableuser=$DisableUserCheckbox.Checked;$objForm.Close()})
$objForm.Controls.Add($OKButton)


#USERNAME LABEL
$UserLabel = New-Object System.Windows.Forms.Label
$UserLabel.Loion = New-Object System.Drawing.Size(10,20) 
$UserLabel.Size = New-Object System.Drawing.Size(425,20) 
$UserLabel.Text = "Username of Terminated Employee"
$objForm.Controls.Add($UserLabel) 

#USERNAME TEXT BOX
$UserTextBox = New-Object System.Windows.Forms.TextBox 
$UserTextBox.Loion = New-Object System.Drawing.Size(10,40) 
$UserTextBox.Size = New-Object System.Drawing.Size(425,20) 
$objForm.Controls.Add($UserTextBox) 

#DISABLE USER CHECKBOX CONTROL
$DisableUserCheckbox = New-Object System.Windows.Forms.Checkbox 
$DisableUserCheckbox.Loion = New-Object System.Drawing.Size(10,80) 
$DisableUserCheckbox.Size = New-Object System.Drawing.Size(425,20)
$DisableUserCheckbox.Text = "Disable The User?"
$objForm.Controls.Add($DisableUserCheckbox)

#Forward Emails to Users Manager
$emailmanagerCheckbox = New-Object System.Windows.Forms.Checkbox 
$emailmanagerCheckbox.Loion = New-Object System.Drawing.Size(10,110) 
$emailmanagerCheckbox.Size = New-Object System.Drawing.Size(425,20)
$emailmanagerCheckbox.Text = "Forward Emails to User's Manager?"
$objForm.Controls.Add($EmailManagercheckbox)

#FORWARD EMAIL LABEL
$FowardEmailLabel = New-Object System.Windows.Forms.Label
$FowardEmailLabel.Loion = New-Object System.Drawing.Size(10,140) 
$FowardEmailLabel.Size = New-Object System.Drawing.Size(425,20)
$FowardEmailLabel.Text = "Forward Emails to someone other than manager? If Yes, Type In Email Address"
$objForm.Controls.Add($FowardEmailLabel)

#FORWARD EMAIL TEXT BOX
$ForwardingTextBox = New-Object System.Windows.Forms.TextBox 
$ForwardingTextBox.Loion = New-Object System.Drawing.Size(10,160) 
$ForwardingTextBox.Size = New-Object System.Drawing.Size(425,20) 
$objForm.Controls.Add($ForwardingTextBox) 

#ENTER TICKET NUMBER TEXT LABEL
$TicketLabel = New-Object System.Windows.Forms.Label
$TicketLabel.Loion = New-Object System.Drawing.Size(10,190) 
$TicketLabel.Size = New-Object System.Drawing.Size(425,20)
$TicketLabel.Text = "Issue Number"
$objForm.Controls.Add($TicketLabel)

$TicketTextBox = New-Object System.Windows.Forms.TextBox 
$TicketTextBox.Loion = New-Object System.Drawing.Size(10,210) 
$TicketTextBox.Size = New-Object System.Drawing.Size(425,250) 
$objForm.Controls.Add($TicketTextBox) 


#CANCEL BUTTONS
$CancelButton = New-Object System.Windows.Forms.Button
$CancelButton.Loion = New-Object System.Drawing.Size(350,320)
$CancelButton.Size = New-Object System.Drawing.Size(75,23)
$CancelButton.Text = "Cancel"
$CancelButton.Add_Click({$objForm.Close(); $cancel = $true})
$objForm.Controls.Add($CancelButton)


$objForm.Topmost = $True
$objForm.Add_Shown({$objForm.Activate()})
[void] $objForm.ShowDialog()
if ($cancel) {return}
#$OKButton.Add_Click({$userinput=$UserTextBox.Text;$ticketnumber=$TicketTextBox.Text;$forwardemail=$ForwardingTextBox.Text;$disableuser=$DisableUserCheckbox.Checked;$objForm.Close()})
#$CancelButton.Add_Click({$objForm.Close()})

#COMMON GLOBAL VARIABLES
$disableusercheckbox=$DisableUserCheckbox.Checked
$userinput=$UserTextBox.Text
$forwardemail=$ForwardingTextBox.Text
$ticketnumber=$TicketTextBox.Text

$Month = Get-Date -format MM
$Day = Get-Date -format dd
$Year = Get-Date -format yyyy




If ($OKButton.Add_Click)  {
    
    
########
#ACTIVE DIRECTORY ACTIONS
#########

#DISABLE THE USER
If ($disableusercheckbox -eq $true)
{
  Disable-ADAccount -Identity $userinput
  $disabled = $userinput + " has been disabled"
} else { 
	$notdisabled = $userinput + " has not been disabled at this time" 
}

#GETS ALL GROUPS USER WAS PART OF BEFORE BLOWING THEM OUT
    $User = $userinput
    $List=@()
    $Groups = Get-ADUser -Identity $User -Properties * | select -ExpandProperty memberof
    foreach($i in $Groups){
    $i = ($i -split ',')[0]
    $List += "`r`n" + ($i -creplace 'CN=|}','')
    }

#GETS THE USERS MANAGER EMAIL ADDRESS
   $User = userinput
   $USERMANAGERACCOUNT = Get-ADUser -identity $User -Properties Manager | Select-Object Name, @{n="ManagerName";e={(Get-ADUser -Identity $_.Manager -Properties emailaddress).emailaddress}}
    
#BLOW OUT GROUPS OF USER EXCEPT DOMAIN USERS
(get-aduser $userinput -properties memberof).memberof|remove-adgroupmember -member $userinput -Confirm:$False
	
#SETS THE USERS TITLE,COMPANY/MANAGER TO DISABLED
set-aduser -identity $userinput -title "contoso  - Disabled $Month/$Day/$Year"
set-aduser -identity $userinput -company $null
set-aduser -identity $userinput -manager $null
set-aduser -identity $userinput -department $null
set-aduser -identity $userinput -description "contoso  - Disabled $Month/$Day/$Year per Issue# $ticketnumber"

#CHANGES THE USERS PASSWORD
$newpwd = ConvertTo-SecureString -String "G00dBye@2O!7$%@!" -AsPlainText �Force
Set-ADAccountPassword $userinput �NewPassword $newpwd -Reset

#MOVES THE USER TO DISABLED USERS
Get-ADUser -Filter { samAccountName -like $userinput } | Move-ADObject �TargetPath "OU=Disabled Users,OU=User Accounts,DC=jacontoso,DC=com"


#HIDES USER FROM GLOBAL ADDRESS BOOK
Set-Mailbox -Identity $userinput -HiddenFromAddressListsEnabled $true

#Forwards to the Manager if checked or to Specified Account per field
If ($emailmanagerCheckbox -eq $true)
{
  Set-Mailbox -Identity $userinput -ForwardingAddress $USERMANAGERACCOUNT
  $emailmanageryes = $userinput + " email account has been forwarded to their manager"
} else { 
  Set-Mailbox -Identity $userinput -ForwardingAddress $forwardemail	
  $emailmanagerno = $userinput + " email account has been forwarded to: (if empty, forwarding is not configured)" 
}

#REMOVES THE SESSION
Remove-PSsession $ExchangeSession 

Start-Sleep -s 2


#STARTS UP OUTLOOK TO FIRE OFF EMAIL
Start-Process Outlook


$Mail = $ol.CreateItem(0)
$Mail.To = "test@jacontoso.com"
$Mail.Subject = "RE: Issue #$ticketnumber Status: AD User Account has been Disabled"
$Mail.Body = @"
Active Directory Changes Per Termination Request:

1. $disabled $notdisabled

2. Password has been set to G00dBye@2O!7$%@!

3. Changed Job Title to:  contoso  - Disabled $Month/$Day/$Year per Issue# $ticketnumber

4. Description, Department, Company, & Manager have all been cleared

5. Removed from the following Groups in AD:
$List

6. Moved to Disabled Users in AD


Exchange Changes per Termination Request

1. Hidden from Global Address Book
2. $forwarded $notforwarded

90 day Status Quo appt has been sent out
"@
$Mail.Send()

}