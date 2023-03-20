    <#
    .SYNOPSIS
    Connect to a Oracle Database server using the supplied Privileged Account Credentials, and change the password for a local account.
    .NOTES
    Requires database connections on in-use Port to be allowed through Firewall, and Oracle Data Access Components to be installed
    #>
    function Set-OraclePassword
    {
    	[CmdletBinding()]
    	param (

    		[String]$UserName,
    		[String]$NewPassword
    	)
    	
    #$SQLScript to be called once a database connection has been established. Add one command per line.
    $SQLScript = @"
    	ALTER USER $UserName IDENTIFIED BY $NewPassword
"@
    	$SQLPORT = 5560
        $PrivilegedAccountPassword = Read-File "C:\PwdReset\oraclepass.txt"
	$ServiceName = 'or'
    	try
    	{
    		Add-Type -Path 'C:\oracle\odp.net\managed\common\Oracle.ManagedDataAccess.dll'
    		
            $HostName = 'http://ebsp.red.hat.local'
    		$SQLConnectionString = "Data Source= (DESCRIPTION =(ADDRESS =(PROTOCOL = TCP)(HOST = " + $HostName + ")(PORT = " + $SQLPORT + "))(CONNECT_DATA =(SERVICE_NAME = " + $ServiceName + ")));User Id=" + $env:USERNAME + ";Password=" + $PrivilegedAccountPassword + ";"
    		$SQLConnectiebsdppsexexon = New-Object Oracle.ManagedDataAccess.Client.OracleConnection($SQLConnectionString)
    		$SQLConnection.Open()
    		$SQLCommand = New-Object Oracle.ManagedDataAccess.Client.OracleCommand($SQLScript, $SQLConnection)
    		$SQLCommand.ExecuteScalar()
    		$SQLConnection.Close()
    		Write-Output "Success"
    	}
    	catch
    	{
    		switch -wildcard ($error[0].Exception.ToString().ToLower())
    		{
    			"*Connect timeout occurred*" { Write-Output "Failed to execute script correctly against Host '$HostName' for the account '$UserName'. Please check the Oracle Host Name and connection properties are correct, and that a firewall is not blocking access."; break }
    			"*logon denied*" { Write-Output "Failed to connect to the Host '$HostName' to reset the password for the account '$UserName'. Please check the Privileged Account Credentials provided are correct."; break }
    			"*user*does not exist*" { Write-Output "Failed to execute script correctly against Host '$HostName' for the account '$UserName'. Error = Account does not exist or you do not have appropriate permissions."; break }
    			"*Unable to resolve*" { Write-Output "Failed to connect to the Host '$HostName' to reset the password for the account '$UserName'. Please check the Oracle Host Name and connection properties are correct."; break }
    			"*TNS: No listener*" { Write-Output "Failed to connect to the Host '$HostName' to reset the password for the account '$UserName'. Please check the Oracle Host Name and connection properties are correct."; break }
    			"*TNS:listener does not currently know*" { Write-Output "Failed to connect to the Host '$HostName' to validate the password for the account '$UserName'. Please check the Oracle Host Name and connection properties are correct."; break }
    			"*Cannot find path*" { Write-Output "Failed to find the Oracle Data Access Components. Either the path specified is incorrect, or the Data Access Components are yet to be installed."; break }
    			"*cannot find the file*" { Write-Output "Failed to find the Oracle Data Access Components. Either the path specified is incorrect, or the Data Access Components are yet to be installed."; break }
    			#Add other wildcard matches here as required
    			default { Write-Output "Failed to reset the password for the account '$UserName' on Host '$HostName'. Error = " + $error[0].Exception }
    		}
    	}
    }
     
    #Make a call to the Set-OraclePassword function
    Set-OraclePassword -HostName '[HostName]' -ServiceName '[ServiceName]' -SQLPort '[DatabasePort]' -Username '[UserName]' -NewPassword '[NewPassword]' -PrivilegedAccountUserName '[PrivilegedAccountUserName]' -PrivilegedAccountPassword '[PrivilegedAccountPassword]'