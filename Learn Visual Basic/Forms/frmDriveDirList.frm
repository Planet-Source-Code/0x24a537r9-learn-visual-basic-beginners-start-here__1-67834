VERSION 5.00
Begin VB.Form frmDriveDirList 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drive - Dir - Listbox"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmDriveDirList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6495
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
      TabIndex        =   9
      Top             =   5760
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmDriveDirList.frx":0ECA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   5835
      Width           =   495
   End
   Begin VB.CommandButton cmdGetFileNamePath 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get File Name and Path..."
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   4440
      Width           =   2175
   End
   Begin VB.CommandButton cmdGetFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get File Name..."
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3960
      Width           =   2175
   End
   Begin VB.CommandButton cmdExtension 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Change Extension..."
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CommandButton cmdGetPathFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Path..."
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdGetPathDir 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Path..."
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2175
   End
   Begin VB.CommandButton cmdGetDrive 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Drive..."
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2430
      Left            =   2520
      TabIndex        =   8
      Top             =   3120
      Width           =   3495
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2115
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   3495
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2520
      TabIndex        =   1
      Top             =   360
      Width           =   3495
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Main"
      ForeColor       =   &H00000000&
      Height          =   5535
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmDriveDirList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdExtension_Click()
    
    File1.Pattern = InputBox("Input a default extension in the format of *.ext1*;*.ext2*;etc.", "Learn Visual Basic")
    
End Sub

Private Sub cmdGetDrive_Click()
    
    MsgBox "Drive indicator 1 is set to the " & Drive1.Drive & " drive.", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub cmdGetFile_Click()
    
    If File1.ListIndex <> -1 Then
        MsgBox "File listbox 1 is set to the file named: " & File1.FileName, vbInformation, "Learn Visual Basic"
    Else
        MsgBox "No file has been selected!", vbExclamation, "Learn Visual Basic"
    End If
    
End Sub

Private Sub cmdGetFileNamePath_Click()
    
    If File1.ListIndex <> -1 Then
        MsgBox "File listbox 1 is set to the file: " & File1.Path & "\" & File1.FileName, vbInformation, "Learn Visual Basic"
    Else
        MsgBox "No file has been selected!", vbExclamation, "Learn Visual Basic"
    End If
    
End Sub

Private Sub cmdGetPathDir_Click()
    
    MsgBox "Directory indicator 1 is set to the path : " & Dir1.Path, vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub cmdGetPathFile_Click()

        MsgBox "File listbox 1 is set to the path: " & File1.Path, vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub Dir1_Change()
    
    File1.Path = Dir1.Path
    
End Sub

Private Sub Drive1_Change()
    
    Dir1.Path = Drive1.Drive
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Drive Dir List"
    
End Sub
