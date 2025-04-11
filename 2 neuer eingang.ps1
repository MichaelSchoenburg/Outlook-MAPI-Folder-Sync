# Mit Array
$Array_NeuerEingang = @()
$global:i = 0
ForEach ($mail in $Folder_NeuerEingang.Items) {
    $Array_NeuerEingang += [EMail]::new($mail.subject, $mail.ReceivedTime, $mail.SenderEmailAddress, $mail.To, $mail.EntryID)
    $global:i++
    Write-Host "$($global:i) von $($Folder_NeuerEingang.Items.Count)"
}
