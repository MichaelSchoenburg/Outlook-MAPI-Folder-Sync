cls

Write-Output "$(Get-Date) - Started Script."
$stopwatch = [system.diagnostics.stopwatch]::StartNew()

# Outlook MAPI initialisieren
Add-Type -assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -comobject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")

$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# Öffentliche Ordner per MAPI initialisieren
$Folder_AlterEingang = $namespace.Folders.Item('Öffentliche Ordner - administrator@Domain.TLD').Folders.Item('Alle Öffentlichen Ordner').Folders.Item("Alter E-Mail-Eingang Stand 11.11.20")
$Folder_AltesArchiv = $namespace.Folders.Item('Öffentliche Ordner - administrator@Domain.TLD').Folders.Item('Alle Öffentlichen Ordner').Folders.Item("Alter E-Mail-Eingang Stand 11.11.20").Folders.Item("Archivierte eMails KUNDE")
$Folder_NeuerEingang = $namespace.Folders.Item('Öffentliche Ordner - administrator@Domain.TLD').Folders.Item('Alle Öffentlichen Ordner').Folders.Item("Steinebach").Folders.Item("E-Mail-Eingang - KUNDE")
$Folder_NeuesArchiv = $namespace.Folders.Item('Öffentliche Ordner - administrator@Domain.TLD').Folders.Item('Alle Öffentlichen Ordner').Folders.Item("Steinebach").Folders.Item("E-Mail-Eingang - KUNDE").Folders.Item("Archivierte E-Mails KUNDE")

$global:currentState = 0
$global:finalState = $Folder_AlterEingang.Items.Count + $Folder_AltesArchiv.Items.Count + $Folder_NeuerEingang.Items.Count + $Folder_NeuesArchiv.Items.Count

$Hash_AlterEingang = @{}
$Hash_AltesArchiv = @{}
$Hash_NeuerEingang = @{}
$Hash_NeuesArchiv = @{}

$ArrayList_Errors = [System.Collections.ArrayList]::new()

function FillUp-Hash ($Hash, $Folder)
{
    ForEach ($mail in $Folder.Items)
    {
        $global:currentState++
        if ($global:currentState % 100 -eq 0) {"$($stopwatch.elapsed.totalseconds) seconds in $($global:currentState) out of $($global:finalState)"}

        try
        {
            $MessageID = $mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F")
            
        }
        catch
        {
            $null = $ArrayList_Errors.Add("No Message ID: EntryID = $($mail.EntryID), Subject = $($mail.subject), SenderEmailAddress = $($mail.SenderEmailAddress), CreationTime = $($mail.CreationTime), ErrorExceptionMessage = $($_.Exception.Message)")
            Write-host "No Message ID" -ForegroundColor red
        }

        try
        {
            $Hash.Add($MessageID, $mail.EntryID)
            
        }
        catch
        {
            $null = $ArrayList_Errors.Add("Error at: Message-ID = $($MessageID) , EntryID = $($mail.EntryID), Subject = $($mail.subject), SenderEmailAddress = $($mail.SenderEmailAddress), CreationTime = $($mail.CreationTime), ErrorExceptionMessage = $($_.Exception.Message)")
            Write-host "Error" -ForegroundColor red
        }
    }
}

FillUp-Hash -Hash $Hash_AlterEingang -Folder $Folder_AlterEingang
FillUp-Hash -Hash $Hash_NeuerEingang -Folder $Folder_NeuerEingang
FillUp-Hash -Hash $Hash_AltesArchiv -Folder $Folder_AltesArchiv
FillUp-Hash -Hash $Hash_NeuesArchiv -Folder $Folder_NeuesArchiv

Write-Output "$(Get-Date) - Finished filling up hashes."

$ArrayList_MissingInNeuerEingang = [System.Collections.ArrayList]::new()
$ArrayList_MissingInNeuesArchiv = [System.Collections.ArrayList]::new()

function Search-Hash ($Hash_SearchObject, $Hash_SearchPool, $ArrayList_MissingIn)
{
    ForEach ($element in $Hash_SearchObject.GetEnumerator())
    {
        $found = $Hash_SearchPool[$element.Name]
        if(-not $found)
        {
            $null = $ArrayList_MissingIn.Add($element.Value)
            
        }
    }
}

Search-Hash -Hash_SearchObject $Hash_AlterEingang -Hash_SearchPool $Hash_NeuerEingang -ArrayList_MissingIn $ArrayList_MissingInNeuerEingang
Search-Hash -Hash_SearchObject $Hash_AltesArchiv -Hash_SearchPool $Hash_NeuesArchiv -ArrayList_MissingIn $ArrayList_MissingInNeuesArchiv

Write-Output "$(Get-Date) - Finished searching hashes."

$stopwatch.stop()
$stopwatch.elapsed
