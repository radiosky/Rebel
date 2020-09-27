Attribute VB_Name = "RebelGlobals"
Public ChatMacro(1, 8) As String
Public MacroFileName As String
Public Sub LoadMacros(MFN$)
'MacroFileName =MFN
On Error Resume Next
If Not IsAFile(MFN$) Then
    SaveMacros (MFN$)

End If
sfnum = FreeFile
Open MFN$ For Input As #sfnum
    For I = 0 To 7
        Line Input #sfnum, A$
        ChatMacro(0, I) = A$
        If EOF(sfnum) Then Exit For
        Line Input #sfnum, A$
        ChatMacro(1, I) = A$
        If EOF(sfnum) Then Exit For
    Next I
getout:
    On Error Resume Next

    Close #sfnum
    Exit Sub



LMerr:
MsgBox "Error Loading " + MFN$ + vbCrLf + Err.Description, vbExclamation, "Rebel Macro Load Error"
Resume getout

End Sub
Public Sub SaveMacros(MFN$)

sfnum = FreeFile
Open MFN$ For Output As #sfnum
For I = 0 To 7
        Print #sfnum, ChatMacro(0, I)
        Print #sfnum, ChatMacro(1, I)
Next I
SaveSetting App.Title, "Settings", "MacroFileName", MFN$
End Sub
