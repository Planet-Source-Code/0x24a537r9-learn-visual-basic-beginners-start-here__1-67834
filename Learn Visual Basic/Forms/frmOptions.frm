VERSION 5.00
Begin VB.Form frmOptions 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Option Buttons"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGetSpecific 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Specific Item Value..."
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   450
      ItemData        =   "frmOptions.frx":08CA
      Left            =   5160
      List            =   "frmOptions.frx":08DA
      TabIndex        =   7
      Top             =   1440
      Width           =   855
   End
   Begin VB.OptionButton btnOption 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hamburger with small fries"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Value           =   -1  'True
      Width           =   3135
   End
   Begin VB.OptionButton btnOption 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Double cheeseburger with a milkshake"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   3135
   End
   Begin VB.OptionButton btnOption 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tater tots and a small Mountain Dew"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   3135
   End
   Begin VB.OptionButton btnOption 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Two apple pies and a milkshake"
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   3135
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
      TabIndex        =   8
      Top             =   2160
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmOptions.frx":08FE
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2235
      Width           =   495
   End
   Begin VB.CommandButton cmdGetDescription 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Item Description..."
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   960
      Width           =   2175
   End
   Begin VB.CommandButton cmdGetNo 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Get Item Number..."
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.Frame fraValues 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Getting Values"
      ForeColor       =   &H00000000&
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdGetDescription_Click()
    
    Dim i As Integer
    
    For i = 0 To 3
        If btnOption(i).Value = True Then
            MsgBox "You have selected: " & btnOption(i).Caption & ".", vbInformation, "Learn Visual Basic"
        End If
    Next
    
End Sub

Private Sub cmdGetNo_Click()
    
    Dim i As Integer
    
    For i = 0 To 3
        If btnOption(i).Value = True Then
            MsgBox "You have selected item # " & i + 1 & ".", vbInformation, "Learn Visual Basic"
        End If
    Next
    
End Sub

Private Sub cmdGetSpecific_Click()
    
    Dim ItemNo As Integer
    
    If List1.ListIndex = -1 Then
        MsgBox "No list item has been selected!", vbExclamation, "Learn Visual Basic"
        Exit Sub
    End If
    ItemNo = List1.ListIndex
    MsgBox "Item " & ItemNo + 1 & "'s value is :" & btnOption(ItemNo).Value, vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Options"
    
End Sub
