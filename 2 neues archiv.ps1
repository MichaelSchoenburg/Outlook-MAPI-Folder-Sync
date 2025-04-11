# Mit Array
cls
$Array_NeuesArchiv = @()
$global:Total = $Folder_NeuesArchiv.Items.Count
$global:i = 0
ForEach ($mail in $Folder_NeuesArchiv.Items) {
    try {
        $Array_NeuesArchiv += [EMail]::new($mail.subject, $mail.ReceivedTime, $mail.SenderEmailAddress, $mail.To, $mail.EntryID)
    } catch {
        Write-Host ""
        Write-Host "Error occured:" -ForegroundColor Red
        Write-Host "Affected object: $($mail.subject), $($mail.ReceivedTime), $($mail.SenderEmailAddress), $($mail.To), $($mail.EntryID)"
        $Error[$Error.Count - 1]
    }
    $global:i++
    Write-Progress -PercentComplete ($global:i/$global:Total*100) -Activity "Processing Mails"
}
