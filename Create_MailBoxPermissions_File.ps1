# variable used to store the filename of the result CSV files
$ResultCsv = "MailboxAccess"

# Create output csv file
New-Item -Path . -Name ".\$($ResultCsv).csv" -ItemType "file" | Out-Null

# variable used to store the all batch folders 
$BatchFolders = Get-ChildItem -Directory -Name

ForEach ($BatchFolder in $BatchFolders) {
    Write-Host $BatchFolder

    $MailboxAccessFiles = Get-Item "$($BatchFolder)\*_MASTER_MailboxAccess.csv"
    $MailboxSendAsFiles = Get-Item "$($BatchFolder)\*_MASTER_MailboxSendAs.csv"
    $MailboxSendOnBehalfFiles = Get-Item "$($BatchFolder)\*_MASTER_MailboxSendOnBehalf.csv"

    ForEach ($MailboxAccessFile in $MailboxAccessFiles) {
        Import-CSV "$($MailboxAccessFile)" | Export-CSV ".\$($ResultCsv).csv" -NoTypeInformation -Append
        Write-Host "$($MailboxAccessFile)"
    }

    ForEach ($MailboxSendAsFile in $MailboxSendAsFiles) {
        Import-CSV "$($MailboxSendAsFile)" | Export-CSV ".\$($ResultCsv).csv" -NoTypeInformation -Append
        Write-Host "$($MailboxSendAsFile)"
    }

    ForEach ($MailboxSendOnBehalfFile in $MailboxSendOnBehalfFiles) {
        Import-CSV "$($MailboxSendOnBehalfFile)" | Export-CSV ".\$($ResultCsv).csv" -NoTypeInformation -Append
        Write-Host "$($MailboxSendOnBehalfFile)"
    }

}



