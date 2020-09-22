VERSION 5.00
Begin VB.Form frmFileManipulation 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Manipulation"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmFileManipulation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMove 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move File..."
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   2175
   End
   Begin VB.CommandButton cmdCheckDrive 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check If Drive Exists..."
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdCheckFolder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check If Folder Exists..."
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   1440
      Width           =   2175
   End
   Begin VB.CommandButton cmdCheckFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Check If File Exists..."
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   1440
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
      TabIndex        =   7
      Top             =   2640
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmFileManipulation.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2715
      Width           =   495
   End
   Begin VB.CommandButton cmdDeleteFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete File..."
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   2175
   End
   Begin VB.CommandButton cmdCreateFolder 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create Folder..."
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdCreateFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Create File..."
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Frame fraFileOperations 
      BackColor       =   &H00FFFFFF&
      Caption         =   "File Operations"
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmFileManipulation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCheckDrive_Click()
    
    On Error GoTo errDat
    ChDrive ("c:\")
    MsgBox "C:\ drive does exist.", vbExclamation, "Learn Visual Basic"
    Exit Sub
    
errDat:
    MsgBox "C:\ drive does not exist.", vbExclamation, "Learn Visual Basic"
End Sub

Private Sub cmdCheckFile_Click()
    
    If FileExist(App.Path & "\Dummy Files\Check File Existence Test.txt") = False Then
        MsgBox "File " & App.Path & "\Dummy Files\Check File Existence Test.txt does not exist.", vbExclamation, "Learn Visual Basic"
    Else
        MsgBox "File " & App.Path & "\Dummy Files\Check File Existence Test.txt does exist.", vbExclamation, "Learn Visual Basic"
    End If
    
End Sub

Private Sub cmdCheckFolder_Click()
    
Dim strPath As String

    strPath = App.Path & "\Dummy Files\Check Folder Existence Test\"
    
    If (Dir(strPath, 16) = "") Then
        MsgBox "Folder " & strPath & " does not exist", vbExclamation, "Learn Visual Basic"
    Else
        MsgBox "Folder " & strPath & " does exist", vbExclamation, "Learn Visual Basic"
    End If
    
End Sub

Private Sub cmdCreateFile_Click()
    
    Open App.Path & "\Dummy Files\Create File Test" For Output As #1
    Close #1
    MsgBox "File created successfully!", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub cmdCreateFolder_Click()
    
    MkDir App.Path & "\Dummy Files\Create Folder Test"
    MsgBox "Folder created successfully!", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub cmdDeleteFile_Click()
    
    If FileExist(App.Path & "\Dummy Files\Delete Files Test") = False Then
        MsgBox "No file to delete!", vbCritical, "Learn Visual Basic"
    Else
        Kill App.Path & "\Dummy Files\Delete File Test"
        MsgBox "File deleted successfully!", vbInformation, "Learn Visual Basic"
    End If
    
End Sub

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdMove_Click()
    
    If FileExist(App.Path & "\Dummy Files\Move File Test") = False Then
        MsgBox "File already moved!", vbCritical, "Learn Visual Basic"
    Else
        FileCopy App.Path & "\Dummy Files\Move File Test", App.Path & "\Dummy Files\Move File Test Folder\Move File Test"
        Kill App.Path & "\Dummy Files\Move File Test"
        MsgBox "File moved successfully!", vbInformation, "Learn Visual Basic"
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "File Manipulation"
    
End Sub

Function FileExist(ByVal szFileName As String) As Boolean
    
    Dim nFileNumber As Integer
    
    On Error Resume Next
    
    nFileNumber = FreeFile
    
    '---try to open file
    Open szFileName For Input As nFileNumber
    
    '---if fails the file does not exist
    If Err.Number <> 0 Then
        FileExist = False
    Else
        FileExist = True
    End If
    
    Close nFileNumber
    
End Function

