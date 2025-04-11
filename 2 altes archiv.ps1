# Mit Array
cls
$Array_AltesArchiv = @()
$global:i = 0
$global:Total = $Folder_AltesArchiv.Items.Count
ForEach ($mail in $Folder_AltesArchiv.Items) {
    try {
        $Array_AltesArchiv += [EMail]::new($mail.subject, $mail.ReceivedTime, $mail.SenderEmailAddress, $mail.To, $mail.EntryID)
    } catch {
        Write-Host ""
        Write-Host "Error occured:" -ForegroundColor Red
        Write-Host "Affected object: $($mail.subject), $($mail.ReceivedTime), $($mail.SenderEmailAddress), $($mail.To), $($mail.EntryID)"
        $Error[$Error.Count - 1]
    }
    $global:i++
    Write-Progress -PercentComplete ($global:i/$global:Total*100) -Activity "Processing Mails"
}
Write-Host "Done." -ForegroundColor Yellow
