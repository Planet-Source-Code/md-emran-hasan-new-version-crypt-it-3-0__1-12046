VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Crypt It 3.0"
   ClientHeight    =   5595
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6735
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   6735
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   195
      Left            =   25
      TabIndex        =   3
      Top             =   5370
      Width           =   4450
      _ExtentX        =   7832
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   4935
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   8705
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmMain.frx":0CCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   6720
      Top             =   5400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatusBar 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5325
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7902
            MinWidth        =   1764
            Text            =   "Select the right encryption method for the right encrypted text"
            TextSave        =   "Select the right encryption method for the right encrypted text"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1799
            MinWidth        =   1764
            Picture         =   "frmMain.frx":0D8E
            TextSave        =   "10/11/00"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            Picture         =   "frmMain.frx":132A
            TextSave        =   "7:39 PM"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   6360
      Top             =   5400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":177E
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1890
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19A2
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AB4
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BC6
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CD8
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DEA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EFC
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":200E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2462
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C16
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":73CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":781E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C72
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":80C6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbarIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Open"
            Object.ToolTipText     =   "Open"
            ImageKey        =   "Open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Save"
            Object.ToolTipText     =   "Save"
            ImageKey        =   "Save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "Cut"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy"
            ImageKey        =   "Copy"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "Paste"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Options"
            Object.ToolTipText     =   "Options"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pword"
            Object.ToolTipText     =   "Change Password"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Enc"
            Object.ToolTipText     =   "Encrypt"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Dec"
            Object.ToolTipText     =   "Decrypt"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "Help"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      Begin VB.Line Line1 
         X1              =   0
         X2              =   6720
         Y1              =   360
         Y2              =   360
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu se 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu ser 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuEnc 
         Caption         =   "&Encrypt"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuDec 
         Caption         =   "&Decrypt"
         Shortcut        =   {F6}
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOption 
         Caption         =   "&Options"
      End
      Begin VB.Menu sd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPchange 
         Caption         =   "Change &Password"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuKey 
         Caption         =   "Shortcut &Keys"
      End
      Begin VB.Menu hd 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gblnIgnoreChange As Boolean
Dim gintIndex As Integer
Dim gstrStack(1000) As String

#Const CASE_SENSITIVE_PASSWORD = False
'Encrypt text
Private Function EncryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Encrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function

'Decrypt text encrypted with EncryptText
Private Function DecryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer
    Dim strBuff As String

#If Not CASE_SENSITIVE_PASSWORD Then

    'Convert password to upper case
    'if not case-sensitive
    strPwd = UCase$(strPwd)

#End If

    'Decrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function


Private Sub Form_Load()
Dim f
Dim meth
f = FreeFile
meth = "Run"
If Dir("C:\WINDOWS\SYSTEM\cryptInf.txt") <> "" Then
Else
Open "C:\WINDOWS\SYSTEM\cryptInf.txt" For Output As f
Print #f, meth
Close #f
End If
frmAbout.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub mnuDec_Click()
Dim m
Dim fNum
Dim pWord
fNum = FreeFile
On Error GoTo err_hand
Open "C:\WINDOWS\SYSTEM\cryptInf.txt" For Input As fNum
Input #fNum, m
Close fNum
 
Open "c:\windows\system\pword.txt" For Input As fNum
    Input #fNum, pWord
Close #fNum

If m = "Run" Then
    While Not ProgressBar1.Value > 99
    ProgressBar1.Value = ProgressBar1.Value + 1
    Wend
    Text1.Text = DecryptText(Text1.Text, pWord)
    ProgressBar1.Value = 0
    ProgressBar1.Refresh
    
ElseIf m = "XOR" Then
    While Not ProgressBar1.Value > 99
    ProgressBar1.Value = ProgressBar1.Value + 1
    Wend
    Text1.Text = Encrypt1(Text1.Text)
    ProgressBar1.Value = 0
    ProgressBar1.Refresh
    
ElseIf m = "64" Then
    While Not ProgressBar1.Value > 99
    ProgressBar1.Value = ProgressBar1.Value + 1
    Wend
    Text1.Text = Base64_Decode(Text1.Text)
    ProgressBar1.Value = 0
    ProgressBar1.Refresh
    
Else
    While Not ProgressBar1.Value > 99
    ProgressBar1.Value = ProgressBar1.Value + 1
    Wend
    Call decrypwd
    ProgressBar1.Value = 0
    ProgressBar1.Refresh
    
End If
Exit Sub
err_hand:
MsgBox "Error # 234 : TOO MANY CHARACTER TO LOAD INTO MEMORY", vbExclamation, "Crypt It 3.0"
ProgressBar1.Value = 0
ProgressBar1.Refresh
End Sub


Private Sub mnuEnc_Click()
Dim m
Dim fNum
Dim pWord
fNum = FreeFile
On Error GoTo err
Open "C:\WINDOWS\SYSTEM\cryptInf.txt" For Input As fNum
Input #fNum, m
Close fNum
 
Open "c:\windows\system\pword.txt" For Input As fNum
    Input #fNum, pWord
Close #fNum

If m = "Run" Then
    
   While Not ProgressBar1.Value > 99
   ProgressBar1.Value = ProgressBar1.Value + 1
   Wend
   Text1.Text = EncryptText(Text1.Text, pWord)
   ProgressBar1.Value = 0
   ProgressBar1.Refresh
ElseIf m = "XOR" Then
    While Not ProgressBar1.Value > 99
    ProgressBar1.Value = ProgressBar1.Value + 1
    Wend
    Text1.Text = Encrypt1(Text1.Text)
    ProgressBar1.Value = 0
    ProgressBar1.Refresh
    
ElseIf m = "64" Then
    While Not ProgressBar1.Value > 99
    ProgressBar1.Value = ProgressBar1.Value + 1
    Wend
    Text1.Text = Base64_Encode(Text1.Text)
    ProgressBar1.Value = 0
    ProgressBar1.Refresh
    
Else
    While Not ProgressBar1.Value > 99
    ProgressBar1.Value = ProgressBar1.Value + 1
    Wend
    Call encrypwd
    ProgressBar1.Value = 0
    ProgressBar1.Refresh
    
End If
Exit Sub
err:
MsgBox "Error # 234 : TOO MANY CHARACTER TO LOAD INTO MEMORY", vbExclamation, "Crypt It 3.0"
ProgressBar1.Value = 0
ProgressBar1.Refresh
End Sub

Private Sub mnuKey_Click()
Form4.Show 1, Me
End Sub

Private Sub mnuOption_Click()
Form1.Show 1
End Sub

Private Sub mnuPchange_Click()
Form3.Show
End Sub

Private Sub mnuSelAll_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text1.SetFocus
End Sub
Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.key
        Case "New"
            Call mnuFileNew_Click
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Cut"
            mnuEditCut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Delete"
            Text1.SelText = ""
        Case "Options"
            mnuOption_Click
        Case "Pword"
            mnuPchange_Click
        Case "Enc"
            mnuEnc_Click
        Case "Dec"
            mnuDec_Click
        Case "Help"
            mnuHelpContents_Click
    End Select
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show 1
End Sub

Private Sub mnuHelpContents_Click()
MsgBox "Crypt It 3.0 is a simple encryption utility. Just use these following instruction : " & vbCrLf & "1.Encryption" & vbCrLf & "Just type the text you want to Encrypt or open it using the Open command in File menu. Then select Encrypt from the Action menu. After the Encryption, you can save the file by the Save command." & vbCrLf & vbCrLf & "2.Decryption" & vbCrLf & "Just open the decrypted file using the Open command in File menu. Then select Decrypt from the Action menu." & vbCrLf & vbCrLf & "3.Options" & vbCrLf & "Select Options from Action menu to select the Encryption Method. Use the Change password command to change your password.", vbInformation, "Crypt It 3.0"
End Sub
Private Sub mnuEditPaste_Click()
Text1.SelText = Clipboard.GetText
Text1.SetFocus
End Sub

Private Sub mnuEditCopy_Click()
Clipboard.Clear
Clipboard.SetText Text1.SelText
Text1.SetFocus
End Sub

Private Sub mnuEditCut_Click()
Clipboard.Clear
Clipboard.SetText Text1.SelText
Text1.SelText = ""
Text1.SetFocus
End Sub

Private Sub mnuEditUndo_Click()

If gintIndex = 0 Then Exit Sub
    gblnIgnoreChange = True
    gintIndex = gintIndex - 1
    On Error Resume Next
    Text1.TextRTF = gstrStack(gintIndex)
    gblnIgnoreChange = False
    Text1.SetFocus
End Sub
Private Sub Text1_Change()
If Not gblnIgnoreChange Then
        gintIndex = gintIndex + 1
        gstrStack(gintIndex) = Text1.TextRTF
    End If
End Sub
Private Sub mnuFileExit_Click()
    Unload Me
End Sub


Private Sub mnuFileSave_Click()
Dim sFile As String
On Error Resume Next
With dlgCommonDialog
    .DialogTitle = "Save As..."
    .CancelError = False
    .Filter = "Text File (*.txt)|*.txt|All Files (*.*)|*.*"
    .ShowSave
    If Len(.FileName) = 0 Then
        Exit Sub
    End If
    sFile = .FileName
End With
Open sFile For Output As #1
Print #1, Text1.Text
Close #1
End Sub

Private Sub mnuFileClose_Click()
Text1.Text = ""
Text1.Locked = True
End Sub

Private Sub mnuFileOpen_Click()
Dim sFile As String
Text1.Locked = False
On Error Resume Next
With dlgCommonDialog
    .DialogTitle = "Open"
    .CancelError = False
    .Filter = "Text File (*.txt)|*.txt|All Files (*.*)|*.*"
    .ShowOpen
    If Len(.FileName) = 0 Then
        Exit Sub
    End If
    sFile = .FileName
End With
Text1.LoadFile sFile
End Sub

Private Sub mnuFileNew_Click()
Text1.Locked = False
Text1.Text = ""
End Sub

Private Sub encrypwd()
Dim ori As String
Dim plen As Integer
Dim i As Integer
Dim tchg As String
Dim fchg As String
Dim asval As Integer
On Error GoTo errorhelp
ori = Text1.Text
plen = Len(ori)
i = 1
fchg = ""

While Not i > plen

tchg = Mid(ori, i, 1)
asval = (Asc(tchg) + 4) * 2 - 4
fchg = fchg + Chr(asval)
i = i + 1
Wend
Text1.Text = fchg
GoTo doneit
errorhelp:
doneit:
End Sub



Private Sub decrypwd()
Dim tchg As String
Dim fchg As String
Dim plen As Integer
Dim i As Integer
Dim asval As Integer
On Error GoTo errorhelp
i = 1
plen = Len(Text1.Text)

While Not i > plen
tchg = Mid(Text1.Text, i, 1)
asval = (Asc(tchg) + 4) / 2 - 4
fchg = fchg + Chr(asval)
i = i + 1
Wend
Text1.Text = fchg
GoTo doneit
errorhelp:
doneit:
End Sub

