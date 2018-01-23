 #Title: AD Info Export and Load
#Author: Curtis Forsyth

#Bring in the needed PowerShell Modules
Import-Module ActiveDirectory

#Database Variables, timeout is in seconds
$SQLServer - "g2-hou-rmm01"
$Database = "AD_Info"
$Table = "Users"
$Timeout = 500

#Set the batch size so we can control how much RAM this process uses
$BatchSize = 2000

#CSV Variables
$CSV_File = "C:\Scripts\AD_Information\AD_Info.csv"
$CSV_Delimiter = ","
$FirstRowColumnNames = $true

#Generate the CSV File - Converts AD Time into readable time format
Get-ADUser -Filter * -Properties * | Select-Object AccountExpirationDate,@{Name="accountExpires";Expression={([datetime]::FromFileTime($_.accountExpires))}},AccountLockoutTime,AllowReversiblePasswordEncryption,@{Name="badPasswordTime";Expression={([datetime]::FromFileTime($_.badPasswordTime))}},badPwdCount,CannotChangePassword,CanonicalName,City,CN,Company,Country,countryCode,Created,Deleted,Department,Description,DisplayName,DistinguishedName,EmailAddress,EmployeeID,EmployeeNumber,Enabled,extensionAttribute14,Fax,GivenName,isDeleted,LastBadPasswordAttempt,LastLogonDate,LockedOut,Manager,MemberOf,Modified,Name,ObjectCategory,ObjectClass,ObjectGUID,objectSid,PasswordExpired,PasswordLastSet,PasswordNeverExpires,PasswordNotRequired,@{Name="pwdLastSet";Expression={([datetime]::FromFileTime($_.pwdLastSet))}},SamAccountName,SID,SIDHistory,UserPrincipalName,whenChanged,whenCreated | Export-CSV C:\Scripts\AD_Information\AD_Info.csv -NoTypeInformation
Write-EventLog -LogName "Application" -Source "G2 AD Transformer" -EventId 2 - -EntryType Information -Message "CSV File exported to C:\Scripts\AD-information\AD_Info.csv"

#Start the timer to track process speeds
$Elapsed = [System.Diagnostics.Stopwatch]::StartNew()
[void][Reflection.Assembly]::LoadWithPartialName("System.Data")
[void][Reflection.Assembly]::LoadWithPartialName("System.Data.SqlClient")

　
#Build the SQLBulkCopy connection
$ConnectionString = "Data Source=$SQLServer;Integrated Security=true;Initial Catalog=$Database;"
$bulkcopy = New-Object System.Data.SqlClient.SqlBulkCopy($ConnectionString, [System.Data.SqlClient.SqlBulkCopyOptions]::TableLock)
$bulkcopy.DestinationTableName = $Table
$bulkcopy.bulkcopyTimeout = $Timeout
$bulkcopy.batchsize = $BatchSize

#Create the datatable and generate the columns
$DataTable = New-Object System.Data.DataTable

#Open the text file on the disk and skip the first row where the column names are located
$TextReader = New-Object System.IO.StreamReader($CSV_File)
$Columns = (Get-Content $CSV_File).Split($CSV_Delimiter)
if ($FirstRowColumnNames -eq $true) { $null = $TextReader.ReadLine()}

#Loop through each column and add them to the datatable
foreach ($Column in $Columns) { $null = $DataTable.Columns.Add() }

#Read in the data, line by line
while (($Line = $TextReader.ReadLine()) -ne $null) {

    #Import the data and then empty the table before it takes up too much RAM, but after there are enough table to make each iteration efficient
    $null = $DataTable.Rows.Add($Line.Split($CSV_Delimiter))
    $i++; (($i % $BatchSize) -eq 0) {
        $bulkcopy.WriteToServer($DataTable)
        $DataTable.Clear()
    }
}

#Cleanup and add in all the rows since the last Clear()
if ($DataTable.Rows.Count -gt 0) {
    $bulkcopy.WriteToServer($DataTable)
    $DataTable.Clear()
}

#Housekeeping to clear connections and RAM
$TextReader.Close(); $TextReader.Dispose()
$bulkcopy.Close(); $bulkcopy.Dispose()
$DataTable.Dispose()

Write-EventLog -LogName "Application" -Source "G2 AD Transformer" -EventId 2 - -EntryType Information -Message "Process complete.  $i rows have been imported into the database.  Total elapsed time is $($Elapsed.Elapsed.ToString())"
 
