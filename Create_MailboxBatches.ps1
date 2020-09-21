# variable used to store the filename of the source CSV file
$SourceCsv = "AllMailBoxes"

# variable used to store the filename of the result CSV files
$ResultCsv = "UserList"

# variable used to store the size of the batch
$BatchSize = 1500

#retrieve all mailboxes in the tenant
$AllMailboxes  = Get-Mailbox -ResultSize unlimited | Select-Object -Property Name,PrimarySmtpAddress

#Export the mailboxes to a CSV file
$AllMailboxes | Export-Csv -Path ".\$($SourceCsv).csv" -NoTypeInformation

# variable used to store the total number of mailboxes
$MailboxCount = $AllMailboxes.count

# variable used to advance the number of the row from which the export starts
$startrow = 0

# counter used in names of resulting CSV files
$counter = 1

# setting the while loop to continue as long as the value of the $startrow variable is smaller than the number of rows in your source CSV file
while ($startrow -lt $MailboxCount)
{

# import of however many rows you want the resulting CSV to contain starting from the $startrow position and export of the imported content to a new file
New-Item -Path ".\" -Name "batch$($counter)" -ItemType "directory" | Out-Null
Import-CSV ".\$($SourceCsv).csv" | select-object -skip $startrow -first $BatchSize | Export-CSV ".\batch$($counter)\$($ResultCsv).csv" -NoClobber -NoTypeInformation
Copy-Item ".\Audit-MailboxPermissions.ps1" -Destination ".\batch$($counter)"

# advancing the number of the row from which the export starts
$startrow += $BatchSize

# incrementing the $counter variable
$counter++

}
