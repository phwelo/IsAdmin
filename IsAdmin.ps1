#Connect to the SQL database ScriptOutput Table AdminQuery
$SqlConnection = New-Object System.Data.SqlClient.SqlConnection

#Write-only access via scriptuser sql account
$SqlConnection.ConnectionString = "Server=as2.levelone.local;Database=ScriptOutput;User Id=scriptuser;Password=scr1pt3dPW;"
$SqlConnection.Open()
$Command = New-Object System.Data.SQLClient.SQLCommand
$Command.Connection = $SqlConnection
$CurDate = Get-Date

#Get variables we'll need to reuse
$username = (gc env:username)
$computername = (gc env:computername)
$ipaddresses = [System.Net.Dns]::GetHostAddresses($computername)|select-object IPAddressToString -expandproperty IPAddressToString

#Check if there is a record already
$Query = "SELECT * FROM [ScriptOutput].[dbo].[AdminQuery] WHERE [CompName] LIKE '$computername'"
$Command2=new-object system.Data.SqlClient.SqlCommand($Query,$SqlConnection)
$Dataset=New-Object system.Data.DataSet
$DataAdapter=New-Object system.Data.SqlClient.SqlDataAdapter($Command2)
$ifexist=($DataAdapter.fill($Dataset))

#Check if current user is a member of Administrators and return 
$group =[ADSI]"WinNT://./Administrators,group" 
$members = @($group.psbase.Invoke("Members")) 
$isadmin = ($members | foreach {$_.GetType().InvokeMember("Name", 'GetProperty', $null, $_, $null)}) -contains [Environment]::UserName

#Check for the folder to see if Mortgage Builder is installed
$mbsearch = test-path -path "c:\program files (x86)\mortgage builder","c:\program files\mortgage builder"
$ismb = ($mbsearch -contains "True")

#If Entry Already exists, only update the IsAdmin property
IF ($ifexist -eq "1")
{
$Command.CommandText = "UPDATE [ScriptOutput].[dbo].[AdminQuery] SET IsAdmin = '$isadmin' WHERE CompName = '$computername'"
$Command.ExecuteNonQuery() | out-null
$Command.CommandText = "UPDATE [ScriptOutput].[dbo].[AdminQuery] SET IsMB = '$ismb' WHERE CompName = '$computername'"
$Command.ExecuteNonQuery() | out-null
$Command.CommandText = "UPDATE [ScriptOutput].[dbo].[AdminQuery] SET Updated = '$curdate' WHERE CompName = '$computername'"
$Command.ExecuteNonQuery() | out-null
}
#If it's a new entry, Insert All Values
ELSE
{
#Actually execute the INSERT to SQL
$Command.CommandText = "INSERT INTO AdminQuery (CompName,UserName,IP,IsAdmin,IsMB,Updated) VALUES ('$computername','$username','$ipaddresses','$isadmin','$ismb','$curdate')"
$Command.ExecuteNonQuery() | out-null
}