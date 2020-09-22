VERSION 5.00
Begin VB.Form frmSavingImages 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Saving Images"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmSavingPictures.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save Image..."
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   5040
      Width           =   1695
   End
   Begin VB.Frame fraSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      ForeColor       =   &H00000000&
      Height          =   5415
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6015
      Begin VB.Image imgSample 
         Height          =   6780
         Left            =   720
         Picture         =   "frmSavingPictures.frx":08CA
         Top             =   240
         Width           =   6750
      End
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmSavingPictures.frx":15A10
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5715
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
      TabIndex        =   1
      Top             =   5640
      Width           =   5295
   End
End
Attribute VB_Name = "frmSavingImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdSave_Click()
    
    SavePicture imgSample.Picture, App.Path & "\Dummy Files\Save Image Test.bmp"
    MsgBox "Image saved successfully!", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub Form_Load()
    
    imgSample.Move 720, 240, 4540, 4540
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Saving Images"
    
End Sub
