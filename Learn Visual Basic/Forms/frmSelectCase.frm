VERSION 5.00
Begin VB.Form frmSelectCase 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Case"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmSelectCase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExecute 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Execute..."
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   4800
      Width           =   2175
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmSelectCase.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5475
      Width           =   495
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Done"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   5295
   End
   Begin VB.Frame fraCode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Code"
      ForeColor       =   &H00000000&
      Height          =   5175
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtExample 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1920
         MaxLength       =   1
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "End Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   240
         TabIndex        =   13
         Top             =   4200
         Width           =   2655
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MsgBox ""intExample <> 1 or 2"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   1200
         TabIndex        =   12
         Top             =   3720
         Width           =   4455
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "intExample = "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Case 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   720
         TabIndex        =   10
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Select Case intExample"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MsgBox ""intExample = 1"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   1200
         TabIndex        =   8
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Case 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   6
         Left            =   720
         TabIndex        =   7
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MsgBox ""intExample = 2"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   8
         Left            =   1200
         TabIndex        =   6
         Top             =   2760
         Width           =   3255
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Case Else:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Index           =   9
         Left            =   720
         TabIndex        =   5
         Top             =   3240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSelectCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdExecute_Click()
    
    Dim intExample As Integer
    
    intExample = Val(txtExample)
    
    Select Case intExample
        Case 1: MsgBox "intExample = 1", vbInformation, "Learn Visual Basic"
        Case 2: MsgBox "intExample = 2", vbInformation, "Learn Visual Basic"
        Case Else: MsgBox "intExample <> 1 or 2", vbInformation, "Learn Visual Basic"
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Select Case"
    
End Sub
