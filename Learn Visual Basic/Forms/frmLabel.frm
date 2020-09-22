VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmLabels 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Labels"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmLabel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBackStyle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change BackStyle..."
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdBorder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Border.."
      Height          =   375
      Left            =   4440
      TabIndex        =   3
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton cmdAlignment 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Alignment..."
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   1575
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
      TabIndex        =   7
      Top             =   2640
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmLabel.frx":058A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2715
      Width           =   495
   End
   Begin VB.Frame fraSample 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sample"
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   240
      TabIndex        =   10
      Top             =   360
      Width           =   2415
      Begin VB.Label lblSample 
         BackColor       =   &H00E0E0E0&
         Caption         =   """I've never had major knee surgery on any other part of my body,""                                  - Winston Bennett,"
         ForeColor       =   &H00000000&
         Height          =   855
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdForeColor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Fore Color..."
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdCaption 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Caption..."
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdBackColor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Back Color..."
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   1920
      Width           =   3255
   End
   Begin VB.CommandButton cmdFont 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Font..."
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   480
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5760
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame fraProperties 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Properties"
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmLabels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAlignment_Click()
    
    If lblSample.Alignment <> 2 Then
        lblSample.Alignment = lblSample.Alignment + 1
    Else
        lblSample.Alignment = 0
    End If
    
End Sub

Private Sub cmdBackColor_Click()
    
    Randomize
    lblSample.BackColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
    
End Sub

Private Sub cmdBackStyle_Click()
    
    If lblSample.BackStyle = 0 Then
        lblSample.BackStyle = 1
    Else
        lblSample.BackStyle = 0
    End If
    
End Sub

Private Sub cmdBorder_Click()
    
    If lblSample.BorderStyle = 0 Then
        lblSample.BorderStyle = 1
    Else
        lblSample.BorderStyle = 0
    End If
    
End Sub

Private Sub cmdCaption_Click()
    
    lblSample.Caption = InputBox("Input a new caption for the sample label:", "Learn Visual Basic")
    
End Sub

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdForeColor_Click()
    
    Randomize
    lblSample.ForeColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
    
End Sub

Private Sub cmdFont_Click()
    
    CD.Flags = &H3 + &H100 + &H4000
    CD.ShowFont
    lblSample.Font = CD.FontName
    lblSample.FontBold = CD.FontBold
    lblSample.FontItalic = CD.FontItalic
    lblSample.FontUnderline = CD.FontUnderline
    lblSample.FontStrikethru = CD.FontStrikethru
    lblSample.FontSize = CD.FontSize
    lblSample.ForeColor = CD.Color
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Labels"
    
End Sub
