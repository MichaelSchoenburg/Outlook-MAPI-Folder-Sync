# Suchen:
cls
$Alt = $Array_AltesArchiv # Hier Source Scope angeben
$Array_NotFoundMails = @()
$Array_FoundMails = @()
$Counter = 0
ForEach ($mail in $Alt) {
    $Counter++
    Write-Progress -PercentComplete ($Counter/$Alt.Count*100) -Activity "Processing Mails"

    $NotFound = $false

    # Write-Host ""
    # Write-Host "Search-Object:"
    # Write-Host "Subject: $($mail.Subject)"
    # Write-Host "ReceivedTime: $($mail.ReceivedTime)"
    # Write-Host "From: $($mail.from)"
    # Write-Host "To: $($mail.to)"

    if ($Array_NeuerEingang.Where{$_.ReceivedTime -eq $mail.ReceivedTime}){
        if ($Array_NeuerEingang.Where{$_.subject -eq $mail.subject}) {
            # Write-Host "Found in NeuerEingang" -ForegroundColor Yellow
            $Array_FoundMails += $mail
        } else {
            # Write-Host "Not found in NeuerEingang" -ForegroundColor Gray
        }
    } else {
        # Write-Host "Not found in NeuerEingang" -ForegroundColor Gray

        if ($Array_NeuesArchiv.Where{$_.ReceivedTime -eq $mail.ReceivedTime}){
            if ($Array_NeuesArchiv.Where{$_.subject -eq $mail.subject}) {
                # Write-Host "Found in NeuesArchiv" -ForegroundColor Yellow
                $Array_FoundMails += $mail
            } else {
            # Write-Host "Not found in NeuesArchiv" -ForegroundColor Gray
            $NotFound = $true
            }
        } else {
            # Write-Host "Not found in NeuesArchiv" -ForegroundColor Gray
            $NotFound = $true
        }
    }

    if ($NotFound) {
        # Write-Host "Mail noted as not been found." -ForegroundColor Red
        $Array_NotFoundMails += $mail
    }
}

# $MeinVergleichsHash = [EMail]::new("➘ Alarmschwelle Arnsberg", "18.11.2020 12:10:30", "inventory.noreply@vega.com", "E-Mail-Eingang - RST", "000")

# $Array_NotFoundMails | Export-Csv -Path "P:\$(Get-Date -Format "dd.MM.yyyy hh.mm.ss")#Array_NotFoundMails.CSV" -NoClobber -Encoding UTF8 -Delimiter ";" -NoTypeInformation
# $Array_FoundMails | Export-Csv -Path "P:\$(Get-Date -Format "dd.MM.yyyy hh.mm.ss")#Array_FoundMails.CSV" -NoClobber -Encoding UTF8 -Delimiter ";" -NoTypeInformation
