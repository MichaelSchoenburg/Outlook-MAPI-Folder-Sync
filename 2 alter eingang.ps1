# Mit Array
$Array_AlterEingang = @()
$global:i = 0
ForEach ($mail in $Folder_AlterEingang.Items) {
    $Array_AlterEingang += [EMail]::new($mail.subject, $mail.ReceivedTime, $mail.SenderEmailAddress, $mail.To, $mail.EntryID)
    $global:i++
    Write-Host "$($global:i) von $($Folder_AlterEingang.Items.Count)"
}
