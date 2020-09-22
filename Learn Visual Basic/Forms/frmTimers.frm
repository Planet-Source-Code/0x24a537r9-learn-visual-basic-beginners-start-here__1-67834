VERSION 5.00
Begin VB.Form frmTimers 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Timers"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmTimers.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
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
      TabIndex        =   3
      Top             =   2640
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmTimers.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2715
      Width           =   495
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stop the Timer..."
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   5655
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start the Timer..."
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   5655
   End
   Begin VB.Timer Timer1 
      Left            =   5880
      Top             =   0
   End
   Begin VB.Frame fraSettings 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Timer Settings"
      ForeColor       =   &H00000000&
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6015
      Begin VB.TextBox txtInterval 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         TabIndex        =   2
         Text            =   "100"
         Top             =   1320
         Width           =   4455
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Timer Interval:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
   End
   Begin VB.Shape LED 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   7  'Diagonal Cross
      Height          =   615
      Index           =   1
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Shape LED 
      BorderColor     =   &H00FF0000&
      FillColor       =   &H00FF0000&
      FillStyle       =   7  'Diagonal Cross
      Height          =   615
      Index           =   0
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   1920
      Width           =   2895
   End
End
Attribute VB_Name = "frmTimers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Timer1.Enabled = False
    Unload Me
    
End Sub

Private Sub cmdStart_Click()
    
    Dim i As Integer
    Dim intValue As Integer
    
    For i = 1 To Len(txtInterval.Text)
        intValue = Asc(Mid(txtInterval.Text, i, 1))
        If intValue > 57 Or intValue < 48 Then
            MsgBox "Please enter a valid number in the textbox", vbCritical, "Learn Visual Basic"
            Exit Sub
        End If
    Next
    
    Timer1.Interval = Val(txtInterval.Text)
    Timer1.Enabled = True
    
End Sub

Private Sub cmdStop_Click()
    
    Timer1.Enabled = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Timers"
    
End Sub

Private Sub Timer1_Timer()
    
    If LED(0).FillColor = &HFF0000 Then
        LED(0).FillColor = &HFF&       'R
        LED(0).BorderColor = &HFF&       'R
        LED(1).FillColor = &HFF0000 'B
        LED(1).BorderColor = &HFF0000 'B
    Else
        LED(1).FillColor = &HFF&       'R
        LED(1).BorderColor = &HFF&       'R
        LED(0).FillColor = &HFF0000 'B
        LED(0).BorderColor = &HFF0000 'B
    End If
    
End Sub
