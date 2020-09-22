VERSION 5.00
Begin VB.Form frmForLoop 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "For Loop"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmForLoop.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExecute 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Execute..."
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   4080
      Width           =   2175
   End
   Begin VB.ListBox lstExample 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1425
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   5535
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmForLoop.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4755
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
      TabIndex        =   2
      Top             =   4680
      Width           =   5295
   End
   Begin VB.Frame fraCode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Code"
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtStart 
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
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtEnd 
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
         Left            =   3840
         MaxLength       =   2
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "10"
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox txtStep 
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
         Left            =   4920
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "1"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "For tempForVariable =       to       Step"
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
         TabIndex        =   8
         Top             =   360
         Width           =   4695
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "List1. AddItem tempForVariable"
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
         Left            =   720
         TabIndex        =   7
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Next"
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
         TabIndex        =   6
         Top             =   1320
         Width           =   4695
      End
   End
   Begin VB.Frame fraExample 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Example"
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   6015
   End
End
Attribute VB_Name = "frmForLoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdExecute_Click()
    
    Dim i As Integer
    
    If txtStep = "0" Or (Val(txtStart) > Val(txtEnd) And Val(txtStep) > 0) Or Not IsNumeric(txtStart) Or Not IsNumeric(txtEnd) Or Not IsNumeric(txtStep) Then
        MsgBox "Invalid inputs.", vbExclamation, "Learn Visual Basic"
        Exit Sub
    End If
    
    lstExample.Clear
    For i = Val(txtStart) To Val(txtEnd) Step Val(txtStep)
        lstExample.AddItem i
    Next
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "For Loop"
    
End Sub
