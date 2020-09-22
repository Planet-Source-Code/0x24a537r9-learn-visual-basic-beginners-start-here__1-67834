VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.Form frmTextToSpeech 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Text to Speech"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmTextToSpeech.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS TTS1 
      Height          =   375
      Left            =   5520
      OleObjectBlob   =   "frmTextToSpeech.frx":058A
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton cmdSpeak 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Speak"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   840
      Width           =   1215
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
      Top             =   1440
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmTextToSpeech.frx":05E2
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1515
      Width           =   495
   End
   Begin VB.Frame fraSpeak 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Speak"
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Width           =   3855
      Begin VB.TextBox txtWords 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   3375
      End
   End
End
Attribute VB_Name = "frmTextToSpeech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSpeak_Click()
    
    TTS1.Speak txtWords.Text
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "Text to Speech", "Text to Speech"
    
End Sub
