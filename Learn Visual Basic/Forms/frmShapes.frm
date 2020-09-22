VERSION 5.00
Begin VB.Form frmShapes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Shapes"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmShapes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
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
      TabIndex        =   7
      Top             =   5520
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmShapes.frx":0ECA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   5595
      Width           =   495
   End
   Begin VB.CommandButton cmdApply 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Apply Settings..."
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4800
      Width           =   5775
   End
   Begin VB.ComboBox cmbBackColor 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmShapes.frx":1794
      Left            =   2640
      List            =   "frmShapes.frx":17A7
      TabIndex        =   5
      Text            =   "White"
      Top             =   4320
      Width           =   3375
   End
   Begin VB.ComboBox cmbFStyle 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmShapes.frx":17CD
      Left            =   2640
      List            =   "frmShapes.frx":17E9
      TabIndex        =   4
      Text            =   "7 - Diagonal Cross"
      Top             =   3960
      Width           =   3375
   End
   Begin VB.ComboBox cmbFColor 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmShapes.frx":187C
      Left            =   2640
      List            =   "frmShapes.frx":188F
      TabIndex        =   3
      Text            =   "Blue"
      Top             =   3600
      Width           =   3375
   End
   Begin VB.ComboBox cmbBStyle 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmShapes.frx":18B5
      Left            =   2640
      List            =   "frmShapes.frx":18CE
      TabIndex        =   2
      Text            =   "1 - Solid"
      Top             =   3240
      Width           =   3375
   End
   Begin VB.ComboBox cmbBColor 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmShapes.frx":1933
      Left            =   2640
      List            =   "frmShapes.frx":1946
      TabIndex        =   1
      Text            =   "Blue"
      Top             =   2880
      Width           =   3375
   End
   Begin VB.ComboBox cmbShape 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "frmShapes.frx":196C
      Left            =   2640
      List            =   "frmShapes.frx":1982
      TabIndex        =   0
      Text            =   "0 - Rectangle"
      Top             =   2520
      Width           =   3375
   End
   Begin VB.Frame fraShape 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shape"
      ForeColor       =   &H00000000&
      Height          =   2055
      Left            =   1800
      TabIndex        =   9
      Top             =   120
      Width           =   2655
      Begin VB.Shape Shape 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FF0000&
         FillColor       =   &H00FF0000&
         FillStyle       =   7  'Diagonal Cross
         Height          =   735
         Left            =   120
         Top             =   600
         Width           =   2415
      End
   End
   Begin VB.Frame fraSettings 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Settings"
      ForeColor       =   &H00000000&
      Height          =   3255
      Left            =   120
      TabIndex        =   8
      Top             =   2160
      Width           =   6015
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Background Color:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fill Style:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fill Color:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   13
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Border Style:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Border Color:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shape:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmShapes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdApply_Click()
    
    Shape.Shape = Val(Mid(cmbShape.Text, 1, 1))
    Shape.BorderColor = FindColor(cmbBColor.Text)
    Shape.BorderStyle = Val(Mid(cmbBStyle.Text, 1, 1))
    Shape.FillColor = FindColor(cmbFColor.Text)
    Shape.FillStyle = Val(Mid(cmbFStyle.Text, 1, 1))
    Shape.BackColor = FindColor(cmbBackColor.Text)
        
End Sub

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()

    cmdApply_Click
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Shapes"
    
End Sub

Public Function FindColor(strColor As String)
    
    Select Case strColor
        
        Case "White"
            FindColor = &HFFFFFF
        
        Case "Blue"
            FindColor = &HFF0000
        
        Case "Magenta"
            FindColor = &HFF00FF
        
        Case "Black"
            FindColor = &H0&
        
        Case "Red"
            FindColor = &HFF&
            
    End Select
    
End Function
