VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCombo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Combo Boxes"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmCombo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbSpecific 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmCombo.frx":06EA
      Left            =   4800
      List            =   "frmCombo.frx":0700
      TabIndex        =   12
      Text            =   "1"
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmdFColor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Font's Color..."
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdBColor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Background Color..."
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Specific Item's Value..."
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdGetSelected 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Selected Item.."
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Remove Item..."
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Item..."
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdChangeText 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Text..."
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   480
      Width           =   2175
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
      Top             =   3480
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmCombo.frx":0716
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3555
      Width           =   495
   End
   Begin VB.Frame fraSample 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sample"
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   3375
      Begin MSForms.ComboBox cmbSample 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   3135
         VariousPropertyBits=   746604571
         BackColor       =   16777215
         ForeColor       =   0
         DisplayStyle    =   3
         Size            =   "5530;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "Sample"
         BorderColor     =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame fraSettings 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Settings"
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbSample_Change()
    
    Dim i As Integer
    
    cmbSpecific.Clear
    
    For i = 1 To cmbSample.ListCount
        cmbSpecific.AddItem i
    Next
    
    If cmbSample.ListCount <> 0 Then
        cmbSpecific.ListIndex = 0
    Else
        cmbSpecific.Text = "-"
    End If
    
End Sub

Private Sub cmdAdd_Click()
    
    Dim tmpString As String
    
    tmpString = InputBox("Input text for a new entry:", "Learn Visual Basic")
    cmbSample.AddItem tmpString
    cmbSample_Change
    
End Sub

Private Sub cmdBColor_Click()
    
    cmbSample.BackColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
    
End Sub

Private Sub cmdChangeText_Click()
    
    cmbSample.Text = InputBox("Input new text for the sample Combo Box:", "Learn Visual Basic")
    
End Sub

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdFColor_Click()
    
    cmbSample.ForeColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
    
End Sub

Private Sub cmdGetSelected_Click()
    
    MsgBox "You have selected: Item # " & cmbSample.ListIndex + 1 & ": " & cmbSample.List(cmbSample.ListIndex), vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub cmdRemove_Click()
    
    If cmbSample.ListIndex <> -1 Then
        cmbSample.RemoveItem cmbSample.ListIndex
        cmbSample.ListIndex = 0
        cmbSample_Change
    End If
    
End Sub

Private Sub Command1_Click()
    
    If Val(cmbSpecific.Text) - 1 = cmbSample.ListIndex Then
        MsgBox "The value for entry # " & cmbSpecific.Text & " is True", vbInformation, "Learn Visual Basic"
    Else
        MsgBox "The value for entry # " & cmbSpecific.Text & " is False", vbInformation, "Learn Visual Basic"
    End If
    
End Sub

Private Sub Form_Load()
    
    cmbSample.AddItem "Hello"
    cmbSample.AddItem "World"
    cmbSample.AddItem "What"
    cmbSample.AddItem "Is"
    cmbSample.AddItem "Your"
    cmbSample.AddItem "Name?"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Combo Box"
    
End Sub
