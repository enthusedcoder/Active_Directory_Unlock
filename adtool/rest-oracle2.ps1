cls
# Ora002.ps1
# Need installation of ODAC1120320Xcopy_x64.zip 
# The 32 bit version also exists

# Load the good assembly
Add-Type -Path "C:\app\ssz\product\12.1.0\client_1\odp.net\managed\common\Oracle.ManagedDataAccess.dll"

# Production connexion string
$compConStr = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=*<Valid Host>*)(PORT=*<Valid Port>*)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=*<SErviceNameofDB>*)));User Id=*<User Id>*;Password=*<Password>*;"

# Connexion
$oraConn= New-Object Oracle.ManagedDataAccess.Client.OracleConnection($compConStr)
$oraConn.Open()

# Requête SQL
$sql1 = @"SELECT col FROM tbl1 
WHERE col1='test'
"@

$command1 = New-Object Oracle.ManagedDataAccess.Client.OracleCommand($sql1,$oraConn)

# Execution
$reader1=$command1.ExecuteReader()

while ($reader1.read())
{
  $reader1["col"]  
}

# Fermeture de la conexion
$reader1.Close()
$oraConn.Close()

Write-Output $retObj