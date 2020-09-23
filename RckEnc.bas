Attribute VB_Name = "Module1"
' Name: RckEncrypt()
' Desc: Encrypt or decrypt a string '
' Parameters: 'varWhat --> string to be encrypted or d ' ecrypted 'blnRemoveSpaces (Optional) --> TRUE rem ' ove spaces ' FALSE convert string as is '
' Example: 'RckEncrypt("Encrypt this") = qZWFMD@@\] ' G 'RckEncrypt("qZWFMD@@\]G") = Encryptthis ' 'RckEncrypt("Encrypt this", False) = sXU ' DOFBB^_E 'RckEncrypt("sXUDOFBB^_E, False) = Encr ' ypt this ' '--------------------------------------- '
Function RckEncrypt(varWhat As Variant, Optional blnRemoveSpaces) As String
On Error GoTo EncError
Dim strResult As String, varWhatChr As Integer
Dim EncryptKey As Integer

If IsNull(varWhat) Then
    GoTo EncJumpOut

End If

If Trim(varWhat) = "" Then
    GoTo EncJumpOut

End If

If IsMissing(blnRemoveSpaces) Then
    blnRemoveSpaces = True

End If

If blnRemoveSpaces Then
    varWhat = Replace(varWhat, " ", "")

End If

EncryptKey = Int(Sqr(Len(varWhat) * 81)) + 23

For varWhatChr = 1 To Len(varWhat)

If (Asc(Mid(varWhat, varWhatChr, 1)) Xor EncryptKey) = 0 Then
    strResult = strResult + Chr(255)

ElseIf Asc(Mid(varWhat, varWhatChr, 1)) = 255 Then
    strResult = strResult + Chr(EncryptKey)
Else
strResult = strResult + Chr(Asc(Mid(varWhat, varWhatChr, 1)) Xor EncryptKey)

End If

Next varWhatChr

EncJumpOut:
RckEncrypt = strResult

EncExit:
Exit Function

EncError: MsgBox Err.Description, vbExclamation + vbOKOnly, "rckEncrypt"
Resume EncExit

End Function


