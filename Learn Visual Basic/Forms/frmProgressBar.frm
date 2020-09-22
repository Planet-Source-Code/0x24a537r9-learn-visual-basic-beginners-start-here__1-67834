VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProgressBar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Progress Bars"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmProgressBar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmProgressBar.frx":058A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3915
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
      TabIndex        =   20
      Top             =   3840
      Width           =   5295
   End
   Begin VB.ComboBox cmbAppearance 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2280
      TabIndex        =   19
      Text            =   "3D"
      Top             =   2010
      Width           =   2175
   End
   Begin VB.ComboBox cmbBorder 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2280
      TabIndex        =   18
      Text            =   "None"
      Top             =   1650
      Width           =   2175
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stop..."
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Start..."
      Height          =   375
      Left            =   4800
      TabIndex        =   16
      Top             =   600
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5640
      Top             =   120
   End
   Begin VB.CommandButton cmdSet 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Set Values..."
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      Top             =   2400
      Width           =   2175
   End
   Begin VB.Frame fraTimer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Timer"
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   4680
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame fraManual 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manual"
      ForeColor       =   &H00000000&
      Height          =   2535
      Left            =   240
      TabIndex        =   5
      Top             =   360
      Width           =   4335
      Begin VB.TextBox txtMin 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         TabIndex        =   9
         Text            =   "0"
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtMax 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Text            =   "100"
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtValue 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Text            =   "75"
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Appearance:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Border Style:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Minumum Value:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Maximum Value:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Value:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   3120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   1085
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Smooth"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   2
      Top             =   2280
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Standard"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   4800
      TabIndex        =   1
      Top             =   1920
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Frame fraScrolling 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Scrolling"
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   4680
      TabIndex        =   4
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Settings"
      ForeColor       =   &H00000000&
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSet_Click()
    
    Dim i As Integer
    Dim intValue As Integer
    Dim intApp As Integer
    Dim intBor As Integer
        
    For i = 1 To Len(txtValue.Text)
        intValue = Asc(Mid(txtValue.Text, i, 1))
        If intValue > 57 Or intValue < 48 Then
            MsgBox "Please enter valid numbers in the textboxes", vbCritical, "Learn Visual Basic"
            Exit Sub
        End If
    Next
    For i = 1 To Len(txtMax.Text)
        intValue = Asc(Mid(txtMax.Text, i, 1))
        If intValue > 57 Or intValue < 48 Then
            MsgBox "Please enter valid numbers in the textboxes", vbCritical, "Learn Visual Basic"
            Exit Sub
        End If
    Next
    For i = 1 To Len(txtMin.Text)
        intValue = Asc(Mid(txtMin.Text, i, 1))
        If intValue > 57 Or intValue < 48 Then
            MsgBox "Please enter valid numbers in the textboxes", vbCritical, "Learn Visual Basic"
            Exit Sub
        End If
    Next
    
    If Val(txtMax.Text) < Val(txtMin.Text) Or Val(txtValue.Text) > Val(txtMax.Text) Or Val(txtValue.Text) < Val(txtMin.Text) Then
        MsgBox "Please enter valid numbers in the textboxes", vbCritical, "Learn Visual Basic"
        Exit Sub
    End If
    
    intApp = 1
    intBor = 1
    
    If cmbAppearance.Text = "Flat" Then intApp = 0
    If cmbBorder.Text = "None" Then intBor = 0
    
    Timer1.Enabled = False
    ProgressBar1.Appearance = intApp
    ProgressBar1.BorderStyle = intBor
    ProgressBar1.Min = Val(txtMin.Text)
    ProgressBar1.Max = Val(txtMax.Text)
    ProgressBar1.Value = Val(txtValue.Text)
    
End Sub

Private Sub cmdStart_Click()
    
    ProgressBar1.Value = 0
    Timer1.Enabled = True
    
End Sub

Private Sub cmdStop_Click()
    
    Timer1.Enabled = False
    
End Sub

Private Sub Form_Load()
    
    cmbBorder.AddItem "None"
    cmbBorder.AddItem "Fixed Single"
    cmbAppearance.AddItem "3D"
    cmbAppearance.AddItem "Flat"
    cmdSet_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub Option1_Click()
    
    ProgressBar1.Scrolling = ccScrollingStandard
    
End Sub

Private Sub Option2_Click()
    
    ProgressBar1.Scrolling = ccScrollingSmooth
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "Progress", "Progress"
    
End Sub

Private Sub Timer1_Timer()
    
    ProgressBar1.Value = ProgressBar1.Value + 1
    If ProgressBar1.Value = 100 Then Timer1.Enabled = False
    
End Sub
