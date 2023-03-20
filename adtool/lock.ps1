<#
Requires -Version 3.0
#Requires -Modules ActiveDirectory, GroupPolicy
I used this file to test the user unlock portion of my solution.  All it does is it attempts to authenticate against
Arbys' main domain controller using a supplied AD account and an invalid password, and will continue to do so until the 
user account gets locked out due to too many authentication attempts.  Then I could test to see if the unlock portion of
the form worked, as for some stupid reason, help desk associates didn't have the permissions to unlock AD user accounts
using powershell, but did have the permissions when the exact same functionality was placed in a Visual Basic script.
#>


if ((([xml](Get-GPOReport -Name "Default Password Policy" -ReportType Xml)).GPO.Computer.ExtensionData.Extension.Account |
            Where-Object name -eq LockoutBadCount).SettingNumber) {
    $Password = ConvertTo-SecureString 'upsydaisy!232' -AsPlainText -Force
    $us = Get-ADUser PAlvarez -Properties SamAccountName, UserPrincipalName, LockedOut
  #      ForEach-Object {
            Do {
                Invoke-Command -ComputerName WPMAD611.Red.Hat.Local {Get-Process
               } -Credential (New-Object System.Management.Automation.PSCredential ($($us.UserPrincipalName), $Password) -ErrorAction SilentlyContinue)
            }
           Until ((Get-ADUser -Identity $us.SamAccountName -Properties LockedOut).LockedOut)
            Write-Output "$($us.SamAccountName) has been locked out"
        #}
}
