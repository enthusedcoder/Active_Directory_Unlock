
function Use-RunAs 
{    
     
    param([Switch]$Check) 
     
    $IsAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()` 
        ).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator") 
         
    if ($Check) { return $IsAdmin }     
 
    if ($MyInvocation.ScriptName -ne "") 
    {  
        if (-not $IsAdmin)  
        {  
            try 
            {  
                $arg = "-file `"$($MyInvocation.ScriptName)`"" 
                Start-Process "$psHome\powershell.exe" -Verb Runas -ArgumentList $arg -ErrorAction 'stop'  
            } 
            catch 
            { 
                Write-Warning "Error - Failed to restart script with runas"  
                break               
            } 
            exit # Quit this session of powershell 
        }  
    }  
    else  
    {  
        Write-Warning "Error - Script must be saved as a .ps1 file first"  
        break  
    }  
}
Function Search-User {
Add-Type -Path (Join-Path -Path (Split-Path $script:MyInvocation.MyCommand.Path) -ChildPath 'bin\CubicOrange.Windows.Forms.ActiveDirectory.dll')

$DialogPicker = New-Object CubicOrange.Windows.Forms.ActiveDirectory.DirectoryObjectPickerDialog

$DialogPicker.AllowedLocations = [CubicOrange.Windows.Forms.ActiveDirectory.Locations]::All
$DialogPicker.AllowedObjectTypes = [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Users
$DialogPicker.DefaultLocations = [CubicOrange.Windows.Forms.ActiveDirectory.Locations]::JoinedDomain
$DialogPicker.DefaultObjectTypes = [CubicOrange.Windows.Forms.ActiveDirectory.ObjectTypes]::Users
$DialogPicker.ShowAdvancedView = $false
$DialogPicker.MultiSelect = $true
$DialogPicker.SkipDomainControllerCheck = $true
$DialogPicker.Providers = [CubicOrange.Windows.Forms.ActiveDirectory.ADsPathsProviders]::Default

$DialogPicker.AttributesToFetch.Add('samAccountName')
#$DialogPicker.AttributesToFetch.Add('title')
#$DialogPicker.AttributesToFetch.Add('department')
#$DialogPicker.AttributesToFetch.Add('distinguishedName')


$DialogPicker.ShowDialog()

return $DialogPicker.Selectedobject
}


Function Obtain-areasup
{
    $salepass = Get-Content -Path "C:\PwdReset\salespass.txt"
    $salestoken = Get-Content -Path "C:\PwdReset\salestoken.txt"
    If (((Get-ADUser -Identity $textbox2.Text -Properties SamAccountName).SamAccountName).substring(1) -as [int] -eq $null)
    {
        $gmtostore = ((Get-ADUser -Identity $textbox2.Text -Properties Description).Description).Substring(6)
        $gmtostore = -join("A","$gmtostore")
    }
    Else
    {
        $gmtostore = (Get-ADUser -Identity $textbox2.Text -Properties SamAccountName).SamAccountName
    }
    Try
    {
        $tofop = (Get-ADUser -Identity $gmtostore -Properties mail -ErrorAction Stop).mail
    }
    Catch
    {
        Return
    }
        $argemail = "Email='$tofop'"
        $salesusern = (Get-ADUser $env:USERNAME -Properties mail).mail
        $salesforce = Connect-Salesforce -User $salesusern -Password $salepass -SecurityToken $salestoken
        $return = Select-Salesforce -Connection $salesforce -Columns 'Area_Manager_Email__c' -Table 'User' -Where "$argemail"
    #SELECT Area_Manager_Email__c FROM User WHERE $argemail;
        Disconnect-Salesforce -Connection $salesforce
        return $return.Area_Manager_Email__c
}
Function Send-outlook
                {

                    Param(
                    [string]$outmail,
                    [string]$sub1,
                    [string]$bod1
                    ) #end param
                
                
                    $olMailItem = 0
                
                    $olApp = new-object -comobject outlook.application
                
                    $NewMail = $olApp.CreateItem($olMailItem)
                
                    $NewMail.Subject = $sub1
                
                    $NewMail.To = $outmail

                    
                
                    $NewMail.Body = $bod1
                
                    $NewMail.Send()
                }
                Function Use-AlternativeEmail
                {
                    Param(
                    [string]$recadd,
                    [string]$sub2,
                    [string]$bod2
                    ) #end param
                
                    $SMTPServer = "webmail1.arbys.com"
                    $port = 25
                    $AuthMethod = "Credential"
                    $domuser = "$env:USERDOMAIN\$env:USERNAME"
                    $usessl = $false
                    $sendnamer = (Get-ADUser $env:USERNAME).Name
                    $sendadd = (Get-ADUser $env:USERNAME -Properties mail).mail
                    $html = $false
                    Send-Email -SMTPServer $SMTPServer -Port $port -UseSSL $usessl -AuthenticationMethod $AuthMethod -Credential $Cred -SenderName $sendnamer -SenderAddress $sendadd -To $recadd -Subject $sub2 -Body $bod2 -HTMLBody $html
                }
Function Test-ADAuthentication {
    param($username,$password)
    Add-Type -AssemblyName System.DirectoryServices.AccountManagement
    $ct = [System.DirectoryServices.AccountManagement.ContextType]::Domain
    $pc = New-Object System.DirectoryServices.AccountManagement.PrincipalContext $ct,$env:USERDOMAIN
    $pc.ValidateCredentials($username,$password)}

$version = $PSVersionTable.PSVersion.Major
$PSScriptRoot = Split-Path -Parent -Path $MyInvocation.MyCommand.Definition
If (!(Test-Path "C:\PwdReset\dotnetver.txt"))
{
Start-Process C:\PwdReset\dotnet.exe -ArgumentList "C:\PwdReset\dotnetver.txt" -Wait
}
$instdotnet = Get-Content "C:\PwdReset\dotnetver.txt"
If ($version -lt 5)
{
$powerPrompt = [System.Windows.Forms.MessageBox]::Show("You did not read the README.  Go take a look at it.")
Exit
}
If (!(Test-Path "$env:ProgramFiles\WindowsPowershell\modules\carbon"))
{
Get-PackageProvider -Name NuGet -ForceBootstrap
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
 Try
{
	Install-module Carbon -Force -ErrorAction Stop
}
Catch
{
	Install-module Carbon -Force -AllowClobber
}
}
If (!(Test-Path "$env:ProgramFiles\WindowsPowershell\modules\PowerShellCookbook"))
{
Get-PackageProvider -Name NuGet -ForceBootstrap
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
Try
{
	Install-module PowerShellCookbook -Force -ErrorAction Stop
}
Catch
{
	Install-module PowerShellCookbook -Force -AllowClobber
}
}
If (!(Test-Path "$env:ProgramFiles\WindowsPowershell\modules\SendEmail"))
{
Get-PackageProvider -Name NuGet -ForceBootstrap
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted
Install-Module SendEmail
}

	[void][reflection.assembly]::Load("System, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Windows.Forms, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Drawing, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
	[void][reflection.assembly]::Load("mscorlib, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Data, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.Xml, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
	[void][reflection.assembly]::Load("System.DirectoryServices, Version=2.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")

If (!(Test-Path -Path "C:\PwdReset\secpasure.txt") -or !(Test-path -Path "C:\PwdReset\oraclepass.txt") )
{
Do
{
$info = Get-Credential -Credential "$env:USERDOMAIN\$env:USERNAME"
If ($info -eq $null)
{
Exit
}
Else
{
$result = Test-ADAuthentication -username $env:USERNAME -password $info.GetNetworkCredential().Password
}
}
While (!($result))
[System.Windows.Forms.MessageBox]::Show('Password will be encrypted and stored in "C:\PwdReset\secpasure.txt".')
ConvertFrom-SecureString -SecureString $info.Password | Out-file "C:\PwdReset\secpasure.txt"
Read-InputBox -Title "Oracle login Password" | Out-file "C:\PwdReset\oraclepass.txt"
#out-file "C:\PwdReset\nonsecpass.txt" -InputObject $info.GetNetworkCredential().Password
}
If (!(Test-path -Path "C:\PwdReset\salespass.txt") -or !(Test-path -Path "C:\PwdReset\salestoken.txt"))
{
    Start-Process Winword.exe -ArgumentList C:\PwdReset\obtain_security_token.docx
    $boxtext = 'In order to take advantage of the salesforce integration that I have embedded in' `
    + 'the application out of necessity, you will need to obtain your salesforce' `
    + 'secuirty token in order to take advantage of the automation.  A word document detailing how to do this' `
    + 'either has already appeared or will soon to explain how to do this.  Click "ok" to continue.'
    [System.Windows.Forms.MessageBox]::Show($boxtext)
    Read-InputBox -Title "Salesforce login Password" | Out-file "C:\PwdReset\salespass.txt"
    Read-InputBox -Title "Salesforce token" | Out-file "C:\PwdReset\salestoken.txt"
}
$domuser = "$env:USERDOMAIN\$env:USERNAME"
$regpass = $Cred.GetNetworkCredential().password
$securstr = Read-File -Path "C:\PwdReset\secpasure.txt"
$gathpass = ConvertTo-SecureString -String $securstr
$Cred = New-Object –TypeName "System.Management.Automation.PSCredential" –ArgumentList $domuser, $gathpass
    $CURRENTPWD = (Get-Content -Path "C:\PwdReset\default_reset_pwd.txt" -ErrorAction SilentlyContinue)
    $CHANGEPWD = [System.Windows.Forms.MessageBox]::Show("Default reset password is '$CURRENTPWD'.  Would you like to change it?","",4)
    if ($CHANGEPWD -eq 'YES')
    {
    C:\PwdReset\default_reset_pwd.txt
    }


	[System.Windows.Forms.Application]::EnableVisualStyles()
	$MainForm = New-Object 'System.Windows.Forms.Form'
	$labelPasswordReset = New-Object 'System.Windows.Forms.Label'
	$textbox2 = New-Object 'System.Windows.Forms.TextBox'
	$textbox1 = New-Object 'System.Windows.Forms.TextBox'
	$buttonUnlockUserAccount = New-Object 'System.Windows.Forms.Button'
	$buttonResetUserPassword = New-Object 'System.Windows.Forms.Button'
    $buttonSendEmailSup = New-Object 'System.Windows.Forms.Button'
    $buttonClose = New-Object 'System.Windows.Forms.Button'
    $buttonSearch = New-Object 'System.Windows.Forms.Button'
    $buttonTruck = New-Object 'System.Windows.Forms.Button'
    $buttonOracleEmail = New-Object 'System.Windows.Forms.Button'
    $buttonAlohaEmail = New-Object 'System.Windows.Forms.Button'
    $buttonAlohaHelpReset = New-Object 'System.Windows.Forms.Button'
    $buttonEmailISP = New-Object 'System.Windows.Forms.Button'
    $ISPedit = New-Object 'System.Windows.Forms.RichTextBox'
    $clearEdit = New-Object 'System.Windows.Forms.Button'
    $tooltip1 = New-Object 'System.Windows.Forms.ToolTip'
	$InitialFormWindowState = New-Object 'System.Windows.Forms.FormWindowState'

	
	$OnLoadFormEvent={
	
	}

$ShowHelp={
    Switch ($this.name) {
        "buttonSearch"  {$tip = "Lets you directly search active directory for user"}
        "buttonUnlockUserAccount" {$tip = "Unlock user account for selected user"}
        "buttonResetUserPassword" {$tip = "Reset user account for selected user"}
        "buttonSendEmailSup" {$tip = "Send email to selected user's area supervisor containing user's new email password"}
        "buttonClose" {$tip = "Send selected user (store or GM account applicable) email concerning coupons"}
        "buttonTruck" {$tip = "Send selected user (store or GM account applicable) email concerning truck schedules"}
        "buttonOracle" {$tip = "Send indicated account (GM account applicable only) or area supervisor of indicated account email with oracle password details"}
        "buttonAloha" {$tip = "Send indicated account (GM account applicable only) or area supervisor of indicated account email with oracle password details"}
        "buttonAlohaHelp" {$tip = "Reset help desk 9991 and 9992 accounts the easy way.  Be sure to have the account loaded in Aloha CFC and already on the Above Store settings tab"}
        "buttonISP" {$tip = "Automatically determine ISP and email them with network equipment status indicated by below input."}
      }
    $tooltip1.SetToolTip($this,$tip)
}

    $buttonSearch_Click=
    {
        $textbox2.Text = $(Search-user).FetchedAttributes
    }
	
	$buttonResetUserPassword_Click=
    {
		$user = $textbox2.Text
			if ([string]::IsNullOrEmpty($user) -eq $false)            
            {
                Function Set-AdUserPwd
                { 
                Param( 
                [string]$user,
                [string]$pwd 
                ) #end param 
	
                $strFilter = "(&(objectCategory=User)(sAMAccountName=$user))"  
                $objDomain = New-Object System.DirectoryServices.DirectoryEntry 
                $objSearcher = New-Object System.DirectoryServices.DirectorySearcher 
                $objSearcher.SearchRoot = $objDomain 
                $objSearcher.PageSize = 1000 
                $objSearcher.Filter = $strFilter 
                $userLDAP = $objSearcher.FindOne() | select-object -ExpandProperty Path 
                if ($userLDAP.Length -gt 0)
                    {
                        $oUser = [adsi]"$userLDAP"
                        $setADUserPwdmsgbox = [System.Windows.Forms.MessageBox]::Show("You have selected $userLDAP. Is this correct?","",4)
                        if ($setADUserPwdmsgbox -eq "YES" ) 
                            {
                            Get-ADUser -Filter {SamACcountName -like $user} -Credential $Cred -ErrorAction SilentlyContinue | Set-ADAccountPassword -NewPassword (ConvertTo-SecureString -AsPlainText $pwd -Force) -Reset -Credential $Cred -ErrorAction SilentlyContinue
                            }
                        else
                            {
                            }
                    }
                    else 
                    {
                    [System.Windows.Forms.MessageBox]::Show("This username does not exist. Please try again.")
                    }
                }


                # CALL FUNCTION
                $NEWPWD = Get-Content -Path "C:\PwdReset\default_reset_pwd.txt"
                if ($NEWPWD.Length -gt 0)
                {
                $Reset_Error = $null
                Set-ADUserPwd -user $user -pwd $NEWPWD
                if ((Get-ADUser -Filter {SamACcountName -like $user} -Properties PasswordLastSet -ErrorVariable Reset_Error -ErrorAction SilentlyContinue -Credential $Cred | Select PasswordLastSet -ExpandProperty PasswordLastSet) -gt (Get-Date).AddMinutes(-1))
                    {
                    [System.Windows.Forms.MessageBox]::Show("$user's password has been reset to $NEWPWD.")
                    }
                else
                    {
                    if ($Reset_Error.Length -gt 0)
                        {
                            [System.Windows.Forms.MessageBox]::Show("There was an error using Active Directory. Are you using an account with proper privileges with RSAT installed?")
                        }
                    [System.Windows.Forms.MessageBox]::Show("Reset aborted.")
                    }
                }
                else
                {
                    [System.Windows.Forms.MessageBox]::Show("ERROR! The default_reset_pwd.txt file is missing.")
                }
            }
            
            else
            {
                [System.Windows.Forms.MessageBox]::Show("The username field is empty.")
            }
    }


	$buttonUnlockUserAccount_Click=
    {
		$user = $textbox2.Text
			if ([string]::IsNullOrEmpty($user) -eq $false)            
            {
                Function Unlock-ADUser
                { 
                Param( 
                [string]$user 
                ) #end param 
	
                $strFilter = "(&(objectCategory=User)(sAMAccountName=$user))"  
                $objDomain = New-Object System.DirectoryServices.DirectoryEntry 
                $objSearcher = New-Object System.DirectoryServices.DirectorySearcher 
                $objSearcher.SearchRoot = $objDomain 
                $objSearcher.PageSize = 1000 
                $objSearcher.Filter = $strFilter 
                $userLDAP = $objSearcher.FindOne() | select-object -ExpandProperty Path 
                if ($userLDAP.Length -gt 0)
                    {
                        $oUser = [adsi]"$userLDAP"
                        $setADUserPwdmsgbox = [System.Windows.Forms.MessageBox]::Show("You have selected $userLDAP. Is this correct?","",4)
                        if ($setADUserPwdmsgbox -eq "YES" ) 
                            {
                            Get-ADUser -Filter {SamACcountName -like $user} -ErrorAction SilentlyContinue | Unlock-ADAccount -ErrorAction SilentlyContinue 
                            $vbname = (Get-ADUser $user -Properties DistinguishedName).DistinguishedName
                            cscript.exe UnlockUserAccount.vbs $vbname
                            #$ouser.psbase.invokeset("AccountDisabled","False") 
                            #$ouser.psbase.CommitChanges()
                            } 
                        else
                            {
                            }
                    }
                    else 
                    {
                    [System.Windows.Forms.MessageBox]::Show("This username does not exist. Please try again.")
                    }
                }
            # CALL FUNCTION
                $Unlock_Error = $null
                if ((Get-ADUser -Filter {SamACcountName -like $user} -Properties LockedOut -ErrorVariable Unlock_Error -ErrorAction SilentlyContinue | Select LockedOut -ExpandProperty LockedOut) -eq $False)
                {
                    [System.Windows.Forms.MessageBox]::Show("$user is already unlocked.")
                }
                else
                {
                    Unlock-ADUser -user $user
                    if ((Get-ADUser -Filter {SamACcountName -like $user} -Properties LockedOut -ErrorVariable Unlock_Error -ErrorAction SilentlyContinue | Select LockedOut -ExpandProperty LockedOut) -eq $False)
                    {
                        [System.Windows.Forms.MessageBox]::Show("$user has been unlocked.")
                    }
                    else
                    {
                    if ($Unlock_Error.Length -gt 0)
                        {
                            [System.Windows.Forms.MessageBox]::Show("There was an error using Active Directory. Are you using an account with proper privileges with RSAT installed?")
                        }
                    [System.Windows.Forms.MessageBox]::Show("Unlock aborted.")
                    }

                }
            }
            else
            {
            [System.Windows.Forms.MessageBox]::Show("The username field is empty.")
            }
    }

    $buttonSendEmailSup_Click=
    {
        $userna = $textbox2.Text
        $insup = Obtain-areasup
        $thesoup = Get-ADUser -Filter {Mail -eq $insup} -Properties GivenName,sAMAccountName
        $strFilter = "(&(objectCategory=User)(sAMAccountName=" + $thesoup.sAMAccountName + "))"  
        $objDomain = New-Object System.DirectoryServices.DirectoryEntry 
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher 
        $objSearcher.SearchRoot = $objDomain 
        $objSearcher.PageSize = 1000 
        $objSearcher.Filter = $strFilter 
        $userLDAP = $objSearcher.FindOne() | select-object -ExpandProperty Path 
        if ($userLDAP.Length -gt 0)
        {
            $oUser = [adsi]"$userLDAP"
            $areaSupBox = [System.Windows.Forms.MessageBox]::Show("You have selected $userLDAP for the area supervisor. Is this correct?","",4)
            if ([string]::IsNullOrEmpty($userna) -eq $false -and $areaSupbox -eq "YES")
            {
                $supname = $thesoup.GivenName
                $endname = (Get-ADUser $userna).GivenName
                $endstore = ((Get-ADUser $userna -Properties Description).Description).Substring(6)
                $stro = Read-File -Path "C:\PwdReset\default_reset_pwd.txt"
                $subject1 = "GM Password Reset Request"
                $body1 = "$supname,`r`n" `
                    + "We recently received a request to reset the email password for $endname at store " `
                    + "$endstore.  We have set the password to " + '"' + "$stro" + '"' + ".  At your " `
                    + "earliest convenience, please provide $endname with this updated information.  We appreciate your " `
                    + "cooperation in this regard.  If the end user still has issues logging in, their account may be " `
                    + "locked as well.  Please have them wait 30 minutes before trying to login again.`r`n`r`nHelp Desk"


                $outlookrunning = (Get-Process | Where-Object { $_.Name -eq "outlook" }).Count -gt 0
                [System.Windows.Forms.MessageBox]::Show("Stop.  Its about to work.")
                If ($outlookrunning)
                {
                    Use-AlternativeEmail -recadd $insup -sub2 $subject1 -bod2 $body1
                    [System.Windows.Forms.MessageBox]::Show("Email sent successfully.")
                }
                Else
                {
                    Send-outlook -outmail $insup -sub1 $subject1 -bod1 $body1
                    [System.Windows.Forms.MessageBox]::Show("Email sent successfully.")
                }
            }
            Else
            {
                [System.Windows.Forms.MessageBox]::Show("The username field is empty or Salesforce not updated.")
            }
        }
        Else
        {
            [System.Windows.Forms.MessageBox]::Show("There was an error.")
        }
    }
    $buttonClose_Click=
    {
        $userna = $textbox2.Text
        $strFilter = "(&(objectCategory=User)(sAMAccountName=$userna))"  
        $objDomain = New-Object System.DirectoryServices.DirectoryEntry 
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher 
        $objSearcher.SearchRoot = $objDomain 
        $objSearcher.PageSize = 1000 
        $objSearcher.Filter = $strFilter 
        $userLDAP = $objSearcher.FindOne() | select-object -ExpandProperty Path 
        if ($userLDAP.Length -gt 0)
        {
            $oUser = [adsi]"$userLDAP"
            $areaSupBox = [System.Windows.Forms.MessageBox]::Show("You have selected $userLDAP. Is this correct?","",4)
            if ([string]::IsNullOrEmpty($userna) -eq $false -and $areaSupbox -eq "YES")
            {
                    $recadd2 = (Get-ADUser $recuser -Properties mail).mail
                    $subject2 = "Issues with coupons"
                    $body2 = "To whom this may concern,`r`n" `
                    + "This is the Arbys help desk.  I am emailing you in regards to a ticket that was " `
                    + "submitted concerning the option for a coupon not displaying on your register.  " `
                    + "In order to get all of the information to accurately resolve this issue, you need " `
                    + "to get in touch with the area supervisor for the store and get them to verify when " `
                    + "the sale in question begins, as many stores mistakenly make the coupons " `
                    + "available before the official start date.  If the area supervisor verifies that " `
                    + "the sale has already begun, then please have him/her email the help desk concerning " `
                    + "this issue, and we can proceed further.`r`n`r`nHelp Desk"  


                $outlookrunning = (Get-Process | Where-Object { $_.Name -eq "outlook" }).Count -gt 0
                If ($outlookrunning)
                {
                    Use-AlternativeEmail -recadd $recadd2 -sub2 $subject2 -bod2 $body2
                    [System.Windows.Forms.MessageBox]::Show("Email sent successfully.")
                }
                Else
                {             
                    Send-outlook -outmail $recadd2 -sub1 $subject2 -bod1 $body2
                    [System.Windows.Forms.MessageBox]::Show("Email sent successfully.")
                }
            }
            Else
            {
                [System.Windows.Forms.MessageBox]::Show("The username field is empty.")
            }
        }
        Else
        {
            [System.Windows.Forms.MessageBox]::Show("There was an error.")
        }

    }

    $buttonTruck_Click=
    {
        $userna = $textbox2.Text
        $strFilter = "(&(objectCategory=User)(sAMAccountName=$userna))"  
        $objDomain = New-Object System.DirectoryServices.DirectoryEntry 
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher 
        $objSearcher.SearchRoot = $objDomain 
        $objSearcher.PageSize = 1000 
        $objSearcher.Filter = $strFilter 
        $userLDAP = $objSearcher.FindOne() | select-object -ExpandProperty Path 
        if ($userLDAP.Length -gt 0)
        {
            $oUser = [adsi]"$userLDAP"
            $areaSupBox = [System.Windows.Forms.MessageBox]::Show("You have selected $userLDAP. Is this correct?","",4)
            if ([string]::IsNullOrEmpty($userna) -eq $false -and $areaSupbox -eq "YES")
            {
                    $recadd3 = (Get-ADUser $recuser -Properties mail).mail
                    $subject3 = "Truck schedule change"
                    $body3 = "To whom this may concern,`r`n" `
                    + "This is the Arbys help desk.  I am emailing you in regards to your ticket " `
                    + "to edit the hours for your orders.  In order for us to process your request, " `
                    + "we will need the following information concerning the orders:`r`n" `
                    + "1. Vendor name `r`n2. Day of week the order is placed `r`n3. Day of week the order is delivered" `
                    + "`r`n4. Time of day the delivery is scheduled to be put away. `r`n" `
                    + "Once we have that information, we will be able to process your request.`r`n`r`nHelp Desk"


                $outlookrunning = (Get-Process | Where-Object { $_.Name -eq "outlook" }).Count -gt 0
                If ($outlookrunning)
                {
                    Use-AlternativeEmail -recadd $recadd3 -sub2 $subject3 -bod2 $body3
                    [System.Windows.Forms.MessageBox]::Show("Email sent successfully.")
                }
                Else
                {             
                    Send-outlook -outmail $recadd3 -sub1 $subject3 -bod1 $body3
                    [System.Windows.Forms.MessageBox]::Show("Email sent successfully.")
                }
            }
            Else
            {
                [System.Windows.Forms.MessageBox]::Show("The username field is empty.")
            }
        }
        Else
        {
            [System.Windows.Forms.MessageBox]::Show("There was an error.")
        }
    }

    $buttonOracleEmail_Click=
    {
        $togm = [System.Windows.Forms.MessageBox]::Show('Do you want to send the reset password to the user indicated in textbox (or corporate user) directly?  Select "no" to send to the area supervisor or would like to input a custom address instead.  Select "yes',"",4)
        $subject4 = "Oracle Password reset"
        If ($togm -eq "YES")
        {
            $user1 = Get-ADUser $textbox2.Text -Properties mail,GivenName
            $usertmail = $user1.mail
            $body4 = $user1.GivenName + ",`r`nThis is the Arbys help desk.  Your " `
            + 'Oracle password has been reset to "Welcome365".  If you have any issues, feel ' `
            + "free to call the help desk for further assistance.`r`n`r`nHelp Desk"
        }
        Else
        {
            
            $gmuser = Get-ADUser $textbox2.Text -Properties GivenName,Description
            $holding = Obtain-areasup
            Try
            {
                $user1 = Get-ADUser -Filter {Mail -eq $holding} -Properties GivenName,mail -ErrorAction Stop
                $usertmail = $user1.mail
                $body4 = $user1.GivenName + ",`r`nThis is the Arbys Help Desk.  " + $gmuser.GivenName + " at " + $gmuser.Description + " " `
                + 'has requested to have their Oracle password reset.  I have reset the password to "Welcome365".  Please ' `
                + "provide this information to " + $gmuser.GivenName + " at your earliest convenience.  We appreciate your " `
                + "cooperation in this regard.`r`n`r`nHelp Desk"
            }
            Catch
            {
                $woops = [System.Windows.Forms.MessageBox]::Show("The user you input os either invalid or does not have an area supervisor assigned to them (most likely meaning they are not a GM).  Would you like to manually input an email address manually?","",4)
                If ($woops -eq "YES")
                {
                    $usertmail = Read-InputBox -Title "Email Address"
                    $body4 = "To whom this may concern,`r`nThis is the Arbys help desk.  We recently reset the Oracle Password " `
                    + "for " + $gmuser.GivenName + " at " + $gmuser.Description + ".  If you are the owner of this account, then " `
                    + "feel free to use it to login to Oracle.  If you are not, please pass the information along to the individual " `
                    + "who requested the password reset.  Thank you.`r`n`r`nHelp Desk"
                }
                Else
                {
                    Return
                }
            }
        }
                $outlookrunning = (Get-Process | Where-Object { $_.Name -eq "outlook" }).Count -gt 0
                If ($outlookrunning)
                {
                    Use-AlternativeEmail -recadd $usertmail -sub2 $subject4 -bod2 $body4
                    [System.Windows.Forms.MessageBox]::Show("Email sent successfully.")
                }
                Else
                {             
                    Send-outlook -outmail $usertmail -sub1 $subject4 -bod1 $body4
                    [System.Windows.Forms.MessageBox]::Show("Email sent successfully.")
                }
    }

    $buttonAlohaEmail_Click=
    {
        $togm2 = [System.Windows.Forms.MessageBox]::Show('Do you want to send the reset password to the GM email?  Select "no" to send to the area supervisor instead.',"",4)
        $subject5 = "Aloha Password reset"
        If ($togm2 -eq "YES")
        {
            $user2 = Get-ADUser $textbox2.Text -Properties mail,GivenName
            $user2mail = $user2.mail
            $body5 = $user2.GivenName + ",`r`nThis is the Arbys help desk.  Your " `
            + 'Aloha password has been reset to "Welcome365".  If you have any issues, feel ' `
            + "free to call the help desk for further assistance.`r`n`r`nHelp Desk"
        }
        Else
        {
            $holding2 = Obtain-areasup
            $gmuser2 = Get-ADUser $textbox2.Text -Properties GivenName,Description
            Try
            {
                If (($gmuser2.SamAccountName).substring(1) -as [int] -ne $null)
                {
                    [System.Windows.Forms.MessageBox]::Show("You cannot use a store account for this operation.  Please select a GM account.")
                    Return
                }
                $user2 = Get-ADUser -Filter {Mail -eq $holding2} -Properties GivenName,mail -ErrorAction Stop
                $user2mail = $user2.mail
                $body5 = $user2.GivenName + ",`r`nThis is the Arbys Help Desk.  " + $gmuser2.GivenName + " at " + $gmuser2.Description + " " `
                + 'has requested to have their Oracle password reset.  I have reset the password to "Welcome365".  Please ' `
                + "provide this information to " + $gmuser2.GivenName + " at your earliest convenience.  We appreciate your " `
                + "cooperation in this regard.`r`n`r`nHelp Desk"
            }
            Catch
            {
                $woops2 = [System.Windows.Forms.MessageBox]::Show("The user you input os either invalid or does not have an area supervisor assigned to them (most likely meaning they are not a GM).  Would you like to manually input an email address manually?","",4)
                If ($woops2 -eq "YES")
                {
                    $user2mail = Read-InputBox -Title "Email Address"
                    $body5 = "To whom this may concern,`r`nThis is the Arbys help desk.  We recently reset the Aloha Password " `
                    + "for " + $gmuser2.GivenName + " at " + $gmuser2.Description + ".  If you are the owner of this account, then " `
                    + "feel free to use it to login to Aloha.  If you are not, please pass the information along to the individual " `
                    + "who requested the password reset.  Thank you.`r`n`r`nHelp Desk"
                }
                Else
                {
                    Return
                }
            }
        }
                $outlookrunning = (Get-Process | Where-Object { $_.Name -eq "outlook" }).Count -gt 0
                If ($outlookrunning)
                {
                    Use-AlternativeEmail -recadd $user2mail -sub2 $subject5 -bod2 $body5
                    [System.Windows.Forms.MessageBox]::Show("Email sent successfully.")
                }
                Else
                {             
                    Send-outlook -outmail $user2mail -sub1 $subject5 -bod1 $body5
                    [System.Windows.Forms.MessageBox]::Show("Email sent successfully.")
                }
    }

    $buttonAlohaHelpReset_Click=
    {
        Start-Process "C:\PwdReset\aloha.exe"
    }

    $buttonEmailISP_Click=
    {
        $storeuse = $textbox2.Text
        $adstore = Get-ADUser $storeuse -Properties StreetAddress,State,PostalCode,City,Description
        Start-Process "C:\PwdReset\library.exe" -ArgumentList "$storeuse" -Wait
        $vendor = Get-Content "$env:APPDATA\$storeuse.txt"
        $ticketnum = Read-InputBox "Arbys incident ticket number"
        $subject6 = $adstore.Description + " internet is down"
        If ($vendor -eq "GRANITE")
        {
            $reci = @('broadbandrepair@granitenet.com;arghd@arbys.com')
        }
        Else
        {
            $reci = @('MS.NOCLeads@gtt.net;arghd@arbys.com')
        }
        $body6 = "To whom this may concern,`r`nArbys " + $adstore.Description + ", " `
        + "located at " + $adstore.StreetAddress + ", " + $adstore.City + ", " + $adstore.State + " " `
        + $adstore.PostalCode + ", has lost their internet connection.  Layer one troubleshooting has " `
        + "been performed, but the internet connection has not recovered.  The Arbys Help Desk ticket " `
        + "number for this incident is $ticketnum, and the light status for both the Sonic Wall and " `
        + "are as follows:`r`n" + $ISPedit.Text + "`r`n`r`nLet us know if you need anything else`r`n`r`n" `
        + "Help Desk"

                    Use-AlternativeEmail -recadd $reci -sub2 $subject6 -bod2 $body6
                    [System.Windows.Forms.MessageBox]::Show("Email sent successfully.")
    }

    $ISPedit_Click=
    {
        If ($ISPedit.Text -eq "Input the light status of networking equipment here.")
        {
            $ISPedit.Text = ""
        }
    }

    $clearEdit_Click=
    {
        $ISPedit.Text = "Input the light status of networking equipment here."
    }



	$Form_StateCorrection_Load=
	{
		$MainForm.WindowState = $InitialFormWindowState
	}
	
	$Form_Cleanup_FormClosed=
	{

		try
		{
            $buttonSearch.remove_Click($buttonSearch_Click)
			$buttonUnlockUserAccount.remove_Click($buttonUnlockUserAccount_Click)
			$buttonResetUserPassword.remove_Click($buttonResetUserPassword_Click)
            $buttonSendEmailSup.remove_Click($buttonSendEmailSup_Click)
            $buttonClose.remove_Click($buttonClose_Click)
            $buttonTruck.remove_Click($buttonTruck_Click)
            $buttonOracleEmail.remove_Click($buttonOracleEmail_Click)
            $buttonAlohaEmail.remove_Click($buttonAlohaEmail_Click)
            $buttonAlohaHelpReset.remove_Click($buttonAlohaHelpReset_Click)
            $buttonEmailISP.remove_Click($buttonEmailISP_Click)
            $ISPedit.remove_Click($ISPedit_Click)
            $clearEdit.remove_Click($clearEdit_Click)
            $buttonEmailISP.remove_MouseHover($ShowHelp)
            $buttonSearch.remove_MouseHover($ShowHelp)
            $buttonUnlockUserAccount.remove_MouseHover($ShowHelp)
            $buttonResetUserPassword.remove_MouseHover($ShowHelp)
            $buttonSendEmailSup.remove_MouseHover($ShowHelp)
            $buttonClose.remove_MouseHover($ShowHelp)
            $buttonTruck.remove_MouseHover($ShowHelp)
            $buttonOracleEmail.remove_MouseHover($ShowHelp)
            $buttonAlohaEmail.remove_MouseHover($ShowHelp)
            $buttonAlohaHelpReset.remove_MouseHover($ShowHelp)
            $buttonEmailISP.remove_MouseHover($ShowHelp)
			$MainForm.remove_Load($OnLoadFormEvent)
			$MainForm.remove_Load($Form_StateCorrection_Load)
			$MainForm.remove_FormClosed($Form_Cleanup_FormClosed)
		}
		catch [Exception]
		{ }
	}

	$MainForm.Controls.Add($labelPasswordReset)
	$MainForm.Controls.Add($textbox2)
	$MainForm.Controls.Add($textbox1)
    $MainForm.Controls.Add($buttonSearch)
	$MainForm.Controls.Add($buttonUnlockUserAccount)
	$MainForm.Controls.Add($buttonResetUserPassword)
    $MainForm.Controls.Add($buttonSendEmailSup)
    $MainForm.Controls.Add($buttonClose)
    $MainForm.Controls.Add($buttonTruck)
    $MainForm.Controls.Add($buttonOracleEmail)
    $MainForm.Controls.Add($buttonAlohaEmail)
    $MainForm.Controls.Add($buttonAlohaHelpReset)
    $MainForm.Controls.Add($buttonEmailISP)
    $MainForm.Controls.Add($ISPedit)
    $MainForm.Controls.Add($clearEdit)
	$MainForm.ClientSize = '380, 430'
	$MainForm.Name = "MainForm"
	$MainForm.StartPosition = 'CenterScreen'
	$MainForm.Text = "Arbys All-in-one assistant"
	$MainForm.add_Load($OnLoadFormEvent)

	$labelPasswordReset.Font = "Tahoma, 9.75pt, style=Bold"
	$labelPasswordReset.Location = '3, 6'
	$labelPasswordReset.Name = "labelPasswordReset"
	$labelPasswordReset.Size = '330, 14'
	$labelPasswordReset.TabIndex = 6
	$labelPasswordReset.Text = "Password Reset"
	$labelPasswordReset.TextAlign = 'TopCenter'

	$textbox2.Location = '61, 23'
	$textbox2.Name = "textbox2"
	$textbox2.Size = '275, 20'
	$textbox2.TabIndex = 8

	$textbox1.BackColor = 'ControlLightLight'
	$textbox1.Enabled = $False
	$textbox1.Location = '4, 23'
	$textbox1.Name = "textbox1"
	$textbox1.ReadOnly = $True
	$textbox1.Size = '61, 20'
	$textbox1.TabIndex = 7
	$textbox1.Text = "Username: "

    $buttonSearch.Font = "Tahoma, 8pt"
    $buttonSearch.Location = '338, 23'
    $buttonSearch.Name = "buttonSearch"
    $buttonSearch.Size = '40, 20'
    $buttonSearch.TabIndex = 10
    $buttonSearch.Text = "...."
    $buttonSearch.UseVisualStyleBackColor = $True
    $buttonSearch.add_MouseHover($ShowHelp)
    $buttonSearch.add_Click($buttonSearch_Click)

	$buttonUnlockUserAccount.Font = "Tahoma, 8pt"
	$buttonUnlockUserAccount.Location = '170, 49'
	$buttonUnlockUserAccount.Name = "buttonUnlockUserAccount"
	$buttonUnlockUserAccount.Size = '165, 22'
	$buttonUnlockUserAccount.TabIndex = 11
	$buttonUnlockUserAccount.Text = "Unlock User Account"
	$buttonUnlockUserAccount.UseVisualStyleBackColor = $True
    $buttonUnlockUserAccount.add_MouseHover($ShowHelp)
	$buttonUnlockUserAccount.add_Click($buttonUnlockUserAccount_Click)

	$buttonResetUserPassword.Font = "Tahoma, 8pt"
	$buttonResetUserPassword.Location = '4, 49'
	$buttonResetUserPassword.Name = "buttonResetUserPassword"
	$buttonResetUserPassword.Size = '165, 22'
	$buttonResetUserPassword.TabIndex = 9
	$buttonResetUserPassword.Text = "Reset User Password"
	$buttonResetUserPassword.UseVisualStyleBackColor = $True
    $buttonResetUserPassword.add_MouseHover($ShowHelp)
	$buttonResetUserPassword.add_Click($buttonResetUserPassword_Click)

    $buttonSendEmailSup.Font = "Tahoma, 8pt"
    $buttonSendEmailSup.Location = '4, 75'
    $buttonSendEmailSup.Name = "buttonSendEmailSup"
    $buttonSendEmailSup.Size = "165, 22"
    $buttonSendEmailSup.TabIndex = 12
    $buttonSendEmailSup.Text = "Send email to area supervisor"
    $buttonSendEmailSup.UseVisualStyleBackColor = $True
    $buttonSendEmailSup.add_MouseHover($ShowHelp)
    $buttonSendEmailSup.add_Click($buttonSendEmailSup_Click)

    $buttonClose.Font = "Tahoma, 8pt"
    $buttonClose.Location = '170, 75'
    $buttonClose.Name = "buttonClose"
    $buttonClose.Size = "165, 22"
    $buttonClose.TabIndex = 13
    $buttonClose.Text = "Coupon email"
    $buttonClose.Enabled = $True
    $buttonClose.UseVisualStyleBackColor = $True
    $buttonClose.add_MouseHover($ShowHelp)
    $buttonClose.add_Click($buttonClose_Click)

    $buttonTruck.Font = "Tahoma, 8pt"
    $buttonTruck.Location = '4, 100'
    $buttonTruck.Name = "buttonTruck"
    $buttonTruck.Size = "165, 22"
    $buttonTruck.TabIndex = 14
    $buttonTruck.Text = "Schedule Email"
    $buttonTruck.Enabled = $True
    $buttonTruck.UseVisualStyleBackColor = $True
    $buttonTruck.add_MouseHover($ShowHelp)
    $buttonTruck.add_Click($buttonTruck_Click)

    $buttonOracleEmail.Font = "Tahoma, 8pt"
    $buttonOracleEmail.Location = '170, 100'
    $buttonOracleEmail.Name = "buttonOracle"
    $buttonOracleEmail.Size = "165, 22"
    $buttonOracleEmail.TabIndex = 15
    $buttonOracleEmail.Text = "Oracle Reset"
    $buttonOracleEmail.Enabled = $True
    $buttonOracleEmail.UseVisualStyleBackColor = $True
    $buttonOracleEmail.add_MouseHover($ShowHelp)
    $buttonOracleEmail.add_Click($buttonOracleEmail_Click)

    $buttonAlohaEmail.Font = "Tahoma, 8pt"
    $buttonAlohaEmail.Location = '4, 125'
    $buttonAlohaEmail.Name = "buttonAloha"
    $buttonAlohaEmail.Size = "165, 22"
    $buttonAlohaEmail.TabIndex = 16
    $buttonAlohaEmail.Text = "Reset Aloha"
    $buttonAlohaEmail.Enabled = $True
    $buttonAlohaEmail.UseVisualStyleBackColor = $True
    $buttonAlohaEmail.add_MouseHover($ShowHelp)
    $buttonAlohaEmail.add_Click($buttonAlohaEmail_Click)

    $buttonAlohaHelpReset.Font = "Tahoma, 8pt"
    $buttonAlohaHelpReset.Location = '170, 125'
    $buttonAlohaHelpReset.Name = "buttonAlohaHelp"
    $buttonAlohaHelpReset.Size = "165, 22"
    $buttonAlohaHelpReset.TabIndex = 17
    $buttonAlohaHelpReset.Text = "Reset 9991/9992"
    $buttonAlohaHelpReset.Enabled = $True
    $buttonAlohaHelpReset.UseVisualStyleBackColor = $True
    $buttonAlohaHelpReset.add_MouseHover($ShowHelp)
    $buttonAlohaHelpReset.add_Click($buttonAlohaHelpReset_Click)

    $buttonEmailISP.Font = "Tahoma, 8pt"
    $buttonEmailISP.Location = '4, 150'
    $buttonEmailISP.Name = "buttonISP"
    $buttonEmailISP.Size = "165, 22"
    $buttonEmailISP.TabIndex = 18
    $buttonEmailISP.Text = "Email store ISP"
    $buttonEmailISP.Enabled = $True
    $buttonEmailISP.UseVisualStyleBackColor = $True
    $buttonEmailISP.add_MouseHover($ShowHelp)
    $buttonEmailISP.add_Click($buttonEmailISP_Click)

    $ISPedit.Font = "Tahoma, 8pt"
    $ISPedit.Location = '4, 180'
    $ISPedit.Name = "equipmentinfo"
    $ISPedit.Size = "320, 200"
    $ISPedit.TabIndex = 19
    $ISPedit.Text = "Input the light status of networking equipment here."
    $ISPedit.Enabled = $True
    $ISPedit.add_Click($ISPedit_Click)

    $clearEdit.Font = "Tahoma, 8pt"
    $clearEdit.Location = '135, 390'
    $clearEdit.Name = "buttonEditReset"
    $clearEdit.Size = "100, 30"
    $clearEdit.TabIndex = 20
    $clearEdit.Text = "Clear Info"
    $clearEdit.Enabled = $True
    $clearEdit.
    $clearEdit.UseVisualStyleBackColor = $True
    $clearEdit.add_Click($clearEdit_Click)


	$InitialFormWindowState = $MainForm.WindowState

	$MainForm.add_Load($Form_StateCorrection_Load)

	$MainForm.add_FormClosed($Form_Cleanup_FormClosed)
	return $MainForm.ShowDialog()
