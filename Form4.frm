VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shortcut Keys"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4725
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   2400
      TabIndex        =   13
      Top             =   120
      Width           =   2175
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   25
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + V"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   24
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   23
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "F5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   22
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "F6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   21
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Decrypt"
         Height          =   345
         Left            =   120
         TabIndex        =   20
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Copy"
         Height          =   345
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Paste"
         Height          =   345
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Select All"
         Height          =   345
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Encrypt"
         Height          =   345
         Left            =   120
         TabIndex        =   16
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   2040
         Width           =   915
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   14
         Top             =   2040
         Width           =   1395
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   12
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "F1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   11
         Top             =   2040
         Width           =   1395
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   10
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + S"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   9
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + O"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   1395
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Open a file"
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1395
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Save a file"
         Height          =   345
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1395
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Undo"
         Height          =   345
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1395
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Cut"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Width           =   1395
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "New File"
         Height          =   345
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Ctrl + N"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1320
         TabIndex        =   1
         Top             =   240
         Width           =   1395
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
