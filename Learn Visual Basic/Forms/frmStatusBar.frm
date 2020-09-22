VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatusBar 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Status Bars"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   Icon            =   "frmStatusBar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   7335
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
      TabIndex        =   4
      Top             =   3480
      Width           =   6375
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   6675
      Picture         =   "frmStatusBar.frx":06EA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3555
      Width           =   495
   End
   Begin MSComctlLib.StatusBar stbAlignment 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3995
            Text            =   "Left"
            TextSave        =   "Left"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   3995
            Text            =   "Center"
            TextSave        =   "Center"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   3995
            Text            =   "Right"
            TextSave        =   "Right"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbBevel 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3995
            Text            =   "Inset"
            TextSave        =   "Inset"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   3995
            Text            =   "No Bevel"
            TextSave        =   "No Bevel"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   3995
            Text            =   "Raised"
            TextSave        =   "Raised"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraBevel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Bevel"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   7095
   End
   Begin VB.Frame fraAlignment 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Alignment"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7095
   End
   Begin MSComctlLib.StatusBar stbStyle 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   1473
            MinWidth        =   18
            Text            =   "Text"
            TextSave        =   "Text"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   1473
            MinWidth        =   18
            Text            =   "Caps"
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            AutoSize        =   1
            Object.Width           =   1473
            MinWidth        =   18
            Text            =   "Num"
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   1473
            MinWidth        =   18
            Text            =   "Ins"
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            AutoSize        =   1
            Object.Width           =   1473
            MinWidth        =   18
            Text            =   "Scrl"
            TextSave        =   "SCRL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   1
            Object.Width           =   1473
            MinWidth        =   18
            TextSave        =   "5:48 PM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   1
            Object.Width           =   1473
            MinWidth        =   18
            TextSave        =   "2/11/2007"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   7
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   1473
            MinWidth        =   18
            Text            =   "Kana"
            TextSave        =   "KANA"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraStyle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Style"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   7095
   End
   Begin MSComctlLib.StatusBar stbAutoSize 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2880
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Manual"
            TextSave        =   "Manual"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8943
            Text            =   "Spring"
            TextSave        =   "Spring"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   18
            Text            =   "Contents"
            TextSave        =   "Contents"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraAutoSize 
      BackColor       =   &H00FFFFFF&
      Caption         =   "AutoSize"
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   7095
   End
End
Attribute VB_Name = "frmStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "Status Bars", "Status Bars"
    
End Sub
