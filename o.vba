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
        MailItem.SaveAs Environ("TEMP") & "\" & Trim(Str(Counter)) & ".msg", olMSG 'указать путь к своей папке для сохранения файлов
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
    Set FileFolder = FSO.getfolder(Environ("TEMP")) 'указать путь к своей папке с сохранёнными файлами
    
    Set oNamespace = Application.GetNamespace("MAPI")
    Set FolderSentMail = oNamespace.GetDefaultFolder(olFolderSentMail) 'или другая папка
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
