VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Crypt It"
   ClientHeight    =   3480
   ClientLeft      =   2550
   ClientTop       =   2985
   ClientWidth     =   4320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   232
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   288
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1200
      Left            =   3840
      Top             =   2160
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   3120
      TabIndex        =   0
      Top             =   3000
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "™"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   7
      Top             =   120
      Width           =   345
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "encrypt data with ease"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   6
      Top             =   690
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Flat #3-A, Aziz Co-Operative Housing Complex, ShahBag, Dhaka-1000, Bangladesh."
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   0
      TabIndex        =   5
      Top             =   3000
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "ehasan@yahoo.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   945
      MouseIcon       =   "About.frx":0CCA
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Tag             =   "mailto:ehasan@yahoo.com?subject=%APP%"
      ToolTipText     =   "Click to send email to me"
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblWeb 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "http://www.emran.koolhost.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   960
      MouseIcon       =   "About.frx":0FD4
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Tag             =   "http://www.emran.koolhost.com"
      ToolTipText     =   "Click to visit my home page on the web"
      Top             =   2340
      Width           =   2775
   End
   Begin VB.Image imgAriad 
      Height          =   480
      Left            =   120
      Picture         =   "About.frx":12DE
      Top             =   240
      Width           =   480
   End
   Begin VB.Label lblApp 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "Crypt It 3.0"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   720
      Left            =   1200
      TabIndex        =   2
      Top             =   0
      Width           =   2610
   End
   Begin VB.Label lblCopy 
      BackColor       =   &H80000005&
      Caption         =   $"About.frx":1FA8
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1140
      Left            =   960
      TabIndex        =   1
      Top             =   1020
      Width           =   3195
      WordWrap        =   -1  'True
   End
   Begin VB.Line lneShad 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   288
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Line lneHigh 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   280
      Y1              =   192
      Y2              =   192
   End
   Begin VB.Shape shpAbout 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Index           =   0
      Left            =   0
      Top             =   2880
      Width           =   4335
   End
   Begin VB.Line lneShad 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   54
      X2              =   54
      Y1              =   216
      Y2              =   -9
   End
   Begin VB.Shape shpAbout 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   3180
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   825
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_HelpID = 2338
Option Explicit
DefInt A-Z

Private Declare Function ShellExecute& Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)

Private Sub cmdOK_Click()
Attribute cmdOK_Click.VB_HelpID = 2341
 Unload Me
End Sub


Private Sub Form_Activate()
Attribute Form_Activate.VB_HelpID = 2444
 Screen.MousePointer = 0
End Sub

Private Sub imgAriad_Click()
Label1.Visible = True
End Sub

Private Sub lblWeb_Click(Index As Integer)
Attribute lblWeb_Click.VB_HelpID = 2343
 Dim Ret As Long
 Dim Cmnd$
 On Error Resume Next
  Cmnd$ = lblWeb(Index).Tag
  Cmnd$ = Replace$(Replace$(Cmnd$, "%APP%", lblApp.Caption), " ", "%20")
  Ret = ShellExecute(hWnd, "Open", Cmnd$, "", "", 5)
  If err Then
   MsgBox Error$ & " (" & err & ")", vbCritical, "Web Access"
  ElseIf Ret <= 32 Then
   MsgBox "Unable to run web document (" & Ret & ")", vbCritical, "Web Access"
  End If
 On Error GoTo 0
End Sub


Private Sub Timer1_Timer()
If Label1.ForeColor = vbBlue Then
    Label1.ForeColor = vbRed
Else
    Label1.ForeColor = vbBlue
End If
End Sub
