Option Explicit

Sub ExportSentMail()
Dim oNamespace As Object
Dim MailItem As Object
Dim FolderSentMail As Object

Dim Counter As Long
 
    Set oNamespace = Application.GetNamespace("MAPI")
    Set FolderSentMail = oNamespace.GetDefaultFolder(olFolderSentMail)
    Counter = 1
    For Each MailItem In FolderSentMail.Items
        MailItem.SaveAs Environ("TEMP") & "\" & Trim(Str(Counter)) & ".msg", olMSG 'óêàçàòü ïóòü ê ñâîåé ïàïêå äëÿ ñîõðàíåíèÿ ôàéëîâ
        Counter = Counter + 1
    Next MailItem
End Sub


Sub importSentMail()
Dim oNamespace As Object
Dim ImportItem As Object
Dim FolderSentMail As Object
Dim FolderImported As Object

Dim FSO As Object
Dim FileFolder As Object
Dim ff As Object
 
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set FileFolder = FSO.getfolder(Environ("TEMP")) 'óêàçàòü ïóòü ê ñâîåé ïàïêå ñ ñîõðàí¸ííûìè ôàéëàìè
    
    Set oNamespace = Application.GetNamespace("MAPI")
    Set FolderSentMail = oNamespace.GetDefaultFolder(olFolderSentMail) 'èëè äðóãàÿ ïàïêà
    On Error Resume Next
        FolderSentMail.Folders.Add ("Imported")
    On Error GoTo 0
    Set FolderImported = FolderSentMail.Folders("Imported")
    For Each ff In FileFolder.Files
        If Right(ff.Name, 4) = ".msg" Then
            Set ImportItem = oNamespace.OpenSharedItem(ff.Path)
            ImportItem.Move FolderImported
        End If
    Next ff
End Sub
