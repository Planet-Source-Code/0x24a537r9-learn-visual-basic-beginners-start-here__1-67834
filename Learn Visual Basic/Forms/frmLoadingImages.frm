VERSION 5.00
Begin VB.Form frmLoadingImages 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loading Images"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmLoadingImages.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load Image..."
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   6840
      Width           =   1695
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
      Top             =   7440
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmLoadingImages.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   7515
      Width           =   495
   End
   Begin VB.Frame fraLoad 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load"
      ForeColor       =   &H00000000&
      Height          =   7215
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6015
      Begin VB.Image imgSample 
         Height          =   6375
         Left            =   900
         Top             =   240
         Width           =   4230
      End
   End
End
Attribute VB_Name = "frmLoadingImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdLoad_Click()
    
    imgSample.Picture = LoadPicture(App.Path & "\Dummy Files\Load Image Test.jpg")
    imgSample.Move 900
    MsgBox "Image loaded successfully!", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub Form_Load()
    
    imgSample.Move 720, 240, 4540, 4540
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Loading Images"
    
End Sub

