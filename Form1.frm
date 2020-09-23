VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4035
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opt64 
      Caption         =   "Base 64 Encryption"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Encryption Method"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.OptionButton optBin 
         Caption         =   "Binary Encryption"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton optRun 
         Caption         =   "Rudimentary Encryption"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2175
      End
      Begin VB.OptionButton optXOR 
         Caption         =   "Simple XOR Encryption"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim fNum
Dim EncMethod
fNum = FreeFile
If optRun.Value = True Then
    EncMethod = "Run"
ElseIf optXOR.Value = True Then
    EncMethod = "XOR"
ElseIf opt64.Value = True Then
    EncMethod = "64"
Else
    EncMethod = "Bin"
End If

Open "C:\WINDOWS\SYSTEM\cryptInf.txt" For Output As fNum
Print #fNum, EncMethod
Close fNum

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim fNum
Dim m
fNum = FreeFile

Open "C:\WINDOWS\SYSTEM\cryptInf.txt" For Input As fNum
Input #fNum, m
Close fNum

If m = "Run" Then
    optRun.Value = True
    optXOR.Value = False
    optBin.Value = False
    opt64.Value = False
ElseIf m = "XOR" Then
    optXOR.Value = True
    optRun.Value = False
    optBin.Value = False
    opt64.Value = False
ElseIf m = "Bin" Then
    optBin.Value = True
    optXOR.Value = False
    optRun.Value = False
    opt64.Value = False
Else
    optBin.Value = False
    optXOR.Value = False
    optRun.Value = False
    opt64.Value = True
End If
    
End Sub

Private Sub opt64_Click()
    optRun.Value = False
    optXOR.Value = False
    optBin.Value = False
    opt64.Value = True
End Sub

Private Sub optBin_Click()
    optRun.Value = False
    optXOR.Value = False
    optBin.Value = True
    opt64.Value = False
End Sub

Private Sub optRun_Click()
    optRun.Value = True
    optXOR.Value = False
    optBin.Value = False
    opt64.Value = False
End Sub

Private Sub optXOR_Click()
    optRun.Value = False
    optXOR.Value = True
    optBin.Value = False
    opt64.Value = False
End Sub
