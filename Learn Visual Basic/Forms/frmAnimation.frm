VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAnimation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Animation"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmAnimation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imlAnimation 
      Left            =   5280
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":0C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":0F72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":128C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":15A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":18C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":1BDA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":1EF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":220E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":2528
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":2842
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAnimation.frx":2B5C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer timAnimation 
      Enabled         =   0   'False
      Interval        =   120
      Left            =   4800
      Top             =   360
   End
   Begin VB.Frame fraAnimation 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Animation"
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   2640
      TabIndex        =   2
      Top             =   120
      Width           =   975
      Begin VB.Image imgAnimation 
         Height          =   480
         Left            =   240
         Picture         =   "frmAnimation.frx":2E76
         Top             =   360
         Width           =   480
      End
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
      Top             =   1200
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmAnimation.frx":3180
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1275
      Width           =   495
   End
End
Attribute VB_Name = "frmAnimation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FrameCount As Integer

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    
    timAnimation.Enabled = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    timAnimation.Enabled = False
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "Animation", "Animation"
    
End Sub

Private Sub timAnimation_Timer()
    
    If FrameCount <> 14 Then
        FrameCount = FrameCount + 1
    Else
        FrameCount = 1
    End If
    
    imgAnimation.Picture = imlAnimation.ListImages.Item(FrameCount).Picture
    
End Sub
