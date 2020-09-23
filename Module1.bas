Attribute VB_Name = "Module1"
Public fMainForm As frmMain


Sub Main()
    Set fMainForm = New frmMain
    fMainForm.Show
End Sub


Public Function CryptNow(ByVal Text As String, ByVal key As String) As String
   For i = 1 To Len(Text)
   a = i Mod Len(key): If a = 0 Then a = Len(key)
   Crypt = Crypt & Chr(Asc(Mid(key, a, 1)) Xor Asc(Mid(Text, i, 1)))
Next i
End Function

Public Function Encrypt1(Text As String) As String
Dim i As Integer

    Value = ""
    For i = 1 To Len(Text)
        If Asc(Mid$(Text, i, 1)) < 128 Then
            Value = Value & Chr$(Asc(Mid$(Text, i, 1)) + 127)
        ElseIf Asc(Mid$(Text, i, 1)) = 128 Then
            Value = Value & Chr$(Asc(Mid$(Text, i, 1)) + 127)
        ElseIf Asc(Mid$(Text, i, 1)) > 128 Then
            Value = Value & Chr$(Asc(Mid$(Text, i, 1)) - 127)
        ElseIf Asc(Mid$(Text, i, 1)) = 255 Then
            Value = Value & Chr$(Asc(Mid$(Text, i, 1)) - 128)
        End If
    Next i

Encrypt1 = Value

End Function

