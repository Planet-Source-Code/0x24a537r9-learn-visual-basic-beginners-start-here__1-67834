VERSION 5.00
Object = "{EEE78583-FE22-11D0-8BEF-0060081841DE}#1.0#0"; "XVoice.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Learn Visual Basic"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File Manipulation"
      Height          =   375
      Index           =   18
      Left            =   6360
      TabIndex        =   25
      Top             =   2040
      Width           =   1455
   End
   Begin ACTIVEVOICEPROJECTLibCtl.DirectSS DirectSS1 
      Height          =   735
      Left            =   5160
      OleObjectBlob   =   "frmMain.frx":08CA
      TabIndex        =   23
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Text to Speech"
      Height          =   375
      Index           =   22
      Left            =   4800
      TabIndex        =   20
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Select Case"
      Height          =   375
      Index           =   21
      Left            =   120
      TabIndex        =   17
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Animation"
      Height          =   375
      Index           =   20
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Labels"
      Height          =   375
      Index           =   19
      Left            =   3240
      TabIndex        =   7
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Loading Images"
      Height          =   375
      Index           =   17
      Left            =   6360
      TabIndex        =   9
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Loading Text"
      Height          =   375
      Index           =   16
      Left            =   7920
      TabIndex        =   10
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Saving Images"
      Height          =   375
      Index           =   15
      Left            =   6360
      TabIndex        =   15
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Saving Text"
      Height          =   375
      Index           =   14
      Left            =   7920
      TabIndex        =   16
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "If ... Then"
      Height          =   375
      Index           =   13
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "For Loop"
      Height          =   375
      Index           =   12
      Left            =   7920
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Status Bars"
      Height          =   375
      Index           =   11
      Left            =   3240
      TabIndex        =   19
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Input Boxes"
      Height          =   375
      Index           =   10
      Left            =   1680
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Message Boxes"
      Height          =   375
      Index           =   9
      Left            =   120
      TabIndex        =   11
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Playing Sounds"
      Height          =   375
      Index           =   8
      Left            =   3240
      TabIndex        =   13
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Combo Boxes"
      Height          =   375
      Index           =   7
      Left            =   3240
      TabIndex        =   2
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Shapes"
      Height          =   375
      Index           =   6
      Left            =   1680
      TabIndex        =   18
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Buttons"
      Height          =   375
      Index           =   5
      Left            =   1680
      TabIndex        =   1
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Progress Bars"
      Height          =   375
      Index           =   4
      Left            =   4800
      TabIndex        =   14
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Option Buttons"
      Height          =   375
      Index           =   3
      Left            =   1680
      TabIndex        =   12
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Timers"
      Height          =   375
      Index           =   2
      Left            =   6360
      TabIndex        =   21
      Top             =   6360
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Drive - Dir - List"
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmdLink 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listboxes"
      Height          =   375
      Index           =   0
      Left            =   4800
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblLink 
      Alignment       =   2  'Center
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   24
      Top             =   2640
      Width           =   735
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      X1              =   -120
      X2              =   9480
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Image imgLink 
      Height          =   630
      Index           =   7
      Left            =   3360
      Picture         =   "frmMain.frx":0922
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   21
      Left            =   600
      Picture         =   "frmMain.frx":316C
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   20
      Left            =   600
      Picture         =   "frmMain.frx":3A36
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   18
      Left            =   6840
      Picture         =   "frmMain.frx":3D40
      Top             =   1440
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   17
      Left            =   6840
      Picture         =   "frmMain.frx":460A
      Top             =   2880
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   16
      Left            =   8400
      Picture         =   "frmMain.frx":4ED4
      Top             =   2880
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   15
      Left            =   6840
      Picture         =   "frmMain.frx":579E
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   14
      Left            =   8400
      Picture         =   "frmMain.frx":6068
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   13
      Left            =   600
      Picture         =   "frmMain.frx":6932
      Top             =   2880
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   12
      Left            =   8400
      Picture         =   "frmMain.frx":71FC
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Learn Visual Basic"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   0
      TabIndex        =   22
      Top             =   120
      Width           =   9495
   End
   Begin VB.Image imgLink 
      Height          =   630
      Index           =   11
      Left            =   3360
      Picture         =   "frmMain.frx":7AC6
      Top             =   5640
      Width           =   1200
   End
   Begin VB.Image imgLink 
      Height          =   645
      Index           =   10
      Left            =   1800
      Picture         =   "frmMain.frx":A268
      Top             =   2760
      Width           =   1185
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   9
      Left            =   600
      Picture         =   "frmMain.frx":CAFA
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   630
      Index           =   0
      Left            =   4920
      Picture         =   "frmMain.frx":D3C4
      Top             =   2760
      Width           =   1200
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   8
      Left            =   3720
      Picture         =   "frmMain.frx":FB66
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   720
      Index           =   6
      Left            =   2040
      Picture         =   "frmMain.frx":10430
      Top             =   5640
      Width           =   720
   End
   Begin VB.Image imgLink 
      Height          =   615
      Index           =   5
      Left            =   2040
      Picture         =   "frmMain.frx":112FA
      Top             =   1320
      Width           =   735
   End
   Begin VB.Image imgLink 
      Height          =   315
      Index           =   4
      Left            =   4860
      Picture         =   "frmMain.frx":12AF0
      Top             =   4320
      Width           =   1260
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   3
      Left            =   2160
      Picture         =   "frmMain.frx":13FDE
      Top             =   4320
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   480
      Index           =   2
      Left            =   6840
      Picture         =   "frmMain.frx":148A8
      Top             =   5760
      Width           =   480
   End
   Begin VB.Image imgLink 
      Height          =   720
      Index           =   1
      Left            =   5160
      Picture         =   "frmMain.frx":15172
      Top             =   1320
      Width           =   720
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type INITCOMMONCONTROLSEX_TYPE
dwSize As Long
dwICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As INITCOMMONCONTROLSEX_TYPE) As Long
Private Const ICC_INTERNET_CLASSES = &H800

Private Sub cmdLink_Click(Index As Integer)
    
    frmMain.Hide
    
    Select Case Index
    
        Case 0
            frmListbox.Show
        Case 1
            frmDriveDirList.Show
        Case 2
            frmTimers.Show
        Case 3
            frmOptions.Show
        Case 4
            frmProgressBar.Show
        Case 5
            frmButtons.Show
        Case 6
            frmShapes.Show
        Case 7
            frmCombo.Show
        Case 8
            frmSound.Show
        Case 9
            frmMessageBoxes.Show
        Case 10
            frmInputBoxes.Show
        Case 11
            frmStatusBar.Show
        Case 12
            frmForLoop.Show
        Case 13
            frmIfThen.Show
        Case 14
            frmSavingText.Show
        Case 15
            frmSavingImages.Show
        Case 16
            frmLoadingText.Show
        Case 17
            frmLoadingImages.Show
        Case 18
            frmFileManipulation.Show
        Case 19
            frmLabels.Show
        Case 20
            frmAnimation.Show
        Case 21
            frmSelectCase.Show
        Case 22
            frmTextToSpeech.Show
        Case Else
            frmMain.Show
        
    End Select
    
End Sub

Private Sub Form_Initialize()
    
    Dim comctls As INITCOMMONCONTROLSEX_TYPE ' identifies the control to register
    Dim RetVal As Long ' generic return value
    
    With comctls
    .dwSize = Len(comctls)
    .dwICC = ICC_INTERNET_CLASSES
    End With
    
    RetVal = InitCommonControlsEx(comctls)
    
End Sub

