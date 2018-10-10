
Set oVersions = Vault.ObjectOperations.GetHistory(ObjVer.ObjID)

Dim currentFilesCount, oldFilesCount

currentFilesCount = Vault.ObjectFileOperations.GetFiles(ObjVer).Count

oldFilesCount = currentFilesCount

dim firstfoundVersion : firstfoundVersion = 0
dim firstDiffStateId : firstDiffStateId = 0

For Each x In oVersions
  Dim oCurrentProps : Set oCurrentProps = Vault.ObjectPropertyOperations.GetProperties(x.ObjVer)
  If oCurrentProps.IndexOf(39) <> -1 Then
    foundStateId = Vault.ObjectPropertyOperations.GetProperty(x.ObjVer,39).TypedValue.GetValueAsLookup().Item
    If foundStateId = stateId Then
    firstfoundVersion = x.ObjVer.Version
    End If
   End If
Next

For Each x In oVersions
  If x.ObjVer.Version = firstfoundVersion Then
     oldFilesCount = Vault.ObjectFileOperations.GetFiles(x.ObjVer).Count
    Exit For
  End If
Next

IF oldFilesCount = currentFilesCount Then
    Err.raise mfscriptcancel , "En az bir dosya yüklemeden devam edilemez.!"
End If

