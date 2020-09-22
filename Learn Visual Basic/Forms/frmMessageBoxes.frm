VERSION 5.00
Begin VB.Form frmMessageBoxes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Message Boxes"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmMessageBoxes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   5
      Top             =   2760
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmMessageBoxes.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2835
      Width           =   495
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generate Message Box..."
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   5775
   End
   Begin VB.ComboBox cmbButtons 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmMessageBoxes.frx":1194
      Left            =   2640
      List            =   "frmMessageBoxes.frx":11AA
      TabIndex        =   3
      Text            =   "OK Only"
      Top             =   1560
      Width           =   3375
   End
   Begin VB.ComboBox cmbStyle 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmMessageBoxes.frx":1209
      Left            =   2640
      List            =   "frmMessageBoxes.frx":121C
      TabIndex        =   2
      Text            =   "(None)"
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Frame fraSettings 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Settings"
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   6015
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set Buttons:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set Style:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set Message:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Set Title:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmMessageBoxes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdGenerate_Click()
    
    Dim varButtons As ButtonConstants
    Dim varButtonStyles As ButtonStyleConstants
    
    Select Case cmbButtons.List(cmbButtons.ListIndex)
        'Case -1
        '    varButtons = vbOKOnly
        Case "OK Only"
            varButtons = vbOKOnly
        Case "OK - Cancel"
            varButtons = vbOKCancel
        Case "Yes - No"
            varButtons = vbYesNo
        Case "Yes - No - Cancel"
            varButtons = vbYesNoCancel
        Case "Retry - Cancel"
            varButtons = vbRetryCancel
        Case "Abort - Retry - Ignore"
            varButtons = vbAbortRetryIgnore
    End Select
    
    Select Case cmbStyle.ListIndex
        Case -1
            MsgBox txtPrompt.Text, varButtons, txtTitle.Text
            Exit Sub
        Case 0
            MsgBox txtPrompt.Text, varButtons, txtTitle.Text
            Exit Sub
        Case 1
            varButtonStyles = vbExclamation
        Case 2
            varButtonStyles = vbInformation
        Case 3
            varButtonStyles = vbQuestion
        Case 4
            varButtonStyles = vbCritical
    End Select

    MsgBox txtPrompt.Text, varButtons + varButtonStyles, txtTitle.Text
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Message Box"
    
End Sub
