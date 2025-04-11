class EMail {
        [string]$Subject
        $ReceivedTime
        [string]$From
        [string]$To
        [string]$EntryID; `
    `
        EMail(
        [string]$Subject,
        $ReceivedTime,
        [string]$From,
        [string]$To,
        [string]$EntryID
        ){
            $this.Subject = $Subject
            $this.ReceivedTime = (Get-Date $ReceivedTime).ToString("dd.MM.yyyy HH:mm:ss")
            $this.From = $From
            $this.To = $To
            $this.EntryID = $EntryID
        }
    }

# Outlook MAPI initialisieren
Add-Type -assembly "Microsoft.Office.Interop.Outlook"
$Outlook = New-Object -comobject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")

$inbox = $namespace.GetDefaultFolder([Microsoft.Office.Interop.Outlook.OlDefaultFolders]::olFolderInbox)

# Öffentliche Ordner per MAPI initialisieren
$Folder_AlterEingang = $namespace.Folders.Item('Öffentliche Ordner - administrator@steinebach-chemikalien.de').Folders.Item('Alle Öffentlichen Ordner').Folders.Item("Alter E-Mail-Eingang Stand 11.11.20")
$Folder_AltesArchiv = $namespace.Folders.Item('Öffentliche Ordner - administrator@steinebach-chemikalien.de').Folders.Item('Alle Öffentlichen Ordner').Folders.Item("Alter E-Mail-Eingang Stand 11.11.20").Folders.Item("Archivierte eMails RST")
$Folder_NeuerEingang = $namespace.Folders.Item('Öffentliche Ordner - administrator@steinebach-chemikalien.de').Folders.Item('Alle Öffentlichen Ordner').Folders.Item("Steinebach").Folders.Item("E-Mail-Eingang - RST")
$Folder_NeuesArchiv = $namespace.Folders.Item('Öffentliche Ordner - administrator@steinebach-chemikalien.de').Folders.Item('Alle Öffentlichen Ordner').Folders.Item("Steinebach").Folders.Item("E-Mail-Eingang - RST").Folders.Item("Archivierte E-Mails RST")

# Variablen initialisieren ## Obsolet
# $Array_StoppedTimes = @()
# $i = 0
# $TotalNum = $Folder_AlterEingang.Items.Count + $Folder_AltesArchiv.Items.Count + $Folder_NeuerEingang.Items.Count + $Folder_NeuesArchiv.Items.Count
