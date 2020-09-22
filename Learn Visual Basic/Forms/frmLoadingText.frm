VERSION 5.00
Begin VB.Form frmLoadingText 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading Text"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmLoadingText.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load Text..."
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2520
      Width           =   1695
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmLoadingText.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3195
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
      Top             =   3120
      Width           =   5295
   End
   Begin VB.Frame fraSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load"
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtSample 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1935
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   360
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmLoadingText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdLoad_Click()
    
    Dim tempLoadVar As String
    
    Open App.Path & "\Dummy Files\Load Text Test.txt" For Input As #1
        Input #1, tempLoadVar
        txtSample.Text = tempLoadVar
    Close #1
    
    MsgBox "Text loaded successfully!", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Loading Text"
    
End Sub


