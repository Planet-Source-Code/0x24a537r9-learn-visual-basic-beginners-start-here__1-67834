VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmButtons 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buttons"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmButtons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   5640
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSample 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sample Button"
      Height          =   615
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdCycle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cycle Mouse Pointers..."
      Height          =   375
      Left            =   4200
      TabIndex        =   13
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdFont 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Font..."
      Height          =   375
      Left            =   4440
      TabIndex        =   10
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdDisable 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Disable..."
      Height          =   375
      Left            =   4440
      TabIndex        =   8
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdEnable 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Enable..."
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddIcon 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Icon..."
      Height          =   375
      Left            =   2760
      TabIndex        =   11
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdDownPic 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Down-Picture..."
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdDisabledPic 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Disabled Picture..."
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdCaption 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Caption..."
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdColor 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Face Color..."
      Height          =   375
      Left            =   2160
      TabIndex        =   14
      Top             =   4440
      Width           =   1935
   End
   Begin VB.Frame fraSample 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sample"
      ForeColor       =   &H00000000&
      Height          =   1455
      Left            =   240
      TabIndex        =   20
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Frame fraProperties 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Properties"
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   120
      TabIndex        =   19
      Top             =   2640
      Width           =   6015
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmButtons.frx":06EA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   5235
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
      TabIndex        =   16
      Top             =   5160
      Width           =   5295
   End
   Begin VB.CommandButton cmdGot 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Got Focus..."
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdLost 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Lost Focus..."
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdRClick 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Right Click..."
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdLClick 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Left Click..."
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdMouseUp 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mouse Up..."
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdMouseDown 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mouse Down..."
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdMouseOver 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mouse Over..."
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Frame fraEvents 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Events"
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmButtons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAddIcon_Click()
    
    cmdSample.Top = 3240
    cmdSample.Height = 975
    cmdSample.Picture = LoadPicture(App.Path & "\Icons\Lego.ico")
    
End Sub

Private Sub cmdCaption_Click()
    
    cmdSample.Caption = InputBox("Input a new caption for the sample button:", "Learn Visual Basic")
    
End Sub

Private Sub cmdColor_Click()
    
    Randomize
    cmdSample.BackColor = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
    
End Sub

Private Sub cmdCycle_Click()
    
    If cmdSample.MousePointer = 7 Then
        cmdSample.MousePointer = 0
    Else
        cmdSample.MousePointer = cmdSample.MousePointer + 1
    End If
    
End Sub

Private Sub cmdDisable_Click()
    
    cmdSample.Enabled = False
    
End Sub

Private Sub cmdDisabledPic_Click()
    
        
    cmdSample.Top = 3240
    cmdSample.Height = 975
    cmdSample.DisabledPicture = LoadPicture(App.Path & "\Icons\Disabled.ico")
    
End Sub

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdDownPic_Click()
    
    cmdSample.Top = 3240
    cmdSample.Height = 975
    cmdSample.DownPicture = LoadPicture(App.Path & "\Icons\Down Picture.ico")
    
End Sub

Private Sub cmdEnable_Click()
    
    cmdSample.Enabled = True
    
End Sub

Private Sub cmdFont_Click()
    
    CD.Flags = &H3 + &H100 + &H4000
    CD.ShowFont
    cmdSample.Font = CD.FontName
    cmdSample.FontBold = CD.FontBold
    cmdSample.FontItalic = CD.FontItalic
    cmdSample.FontUnderline = CD.FontUnderline
    cmdSample.FontStrikethru = CD.FontStrikethru
    cmdSample.FontSize = CD.FontSize
    
End Sub

Private Sub cmdGot_GotFocus()
    
    MsgBox "Got focus event triggered.", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub cmdLClick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbLeftButton Then MsgBox "Left click event triggered.", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub cmdLost_LostFocus()
    
    MsgBox "Lost focus event triggered.", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub cmdMouseDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MsgBox "Mouse down event triggered.", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub cmdMouseOver_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MsgBox "Mouse over event triggered.", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub cmdMouseUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    MsgBox "Mouse up event triggered.", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub cmdRClick_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = vbRightButton Then MsgBox "Right click event triggered.", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "Buttons", "Buttons"
    
End Sub
