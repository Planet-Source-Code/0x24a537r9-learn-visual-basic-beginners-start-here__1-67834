VERSION 5.00
Begin VB.Form frmInputBoxes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Boxes"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmInputBoxes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3015
   ScaleMode       =   0  'User
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generate Input Box..."
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   5775
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
      TabIndex        =   4
      Top             =   2280
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmInputBoxes.frx":058A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2355
      Width           =   495
   End
   Begin VB.Frame fraSettings 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Settings"
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtValue 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         TabIndex        =   2
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox txtTitle 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         TabIndex        =   0
         Text            =   "Title"
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox txtPrompt 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Text            =   "Message"
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set Title:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set Message:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set Default Value:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmInputBoxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdGenerate_Click()
    
    InputBox txtPrompt.Text, txtTitle.Text, txtValue.Text
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Input Box"
    
End Sub
