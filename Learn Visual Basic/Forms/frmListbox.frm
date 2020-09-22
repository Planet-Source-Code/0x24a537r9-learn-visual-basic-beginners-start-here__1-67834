VERSION 5.00
Begin VB.Form frmListbox 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listboxes"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmListbox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   6255
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5600
      Picture         =   "frmListbox.frx":06EA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6680
      Width           =   495
   End
   Begin VB.CommandButton cmdClear2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear Listbox..."
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   3840
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Load Items into Listbox..."
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save Items from Listbox..."
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton cmdAddOver 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add All Items to Textbox..."
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   2175
   End
   Begin VB.CommandButton cmdMoveOver 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Move Selected Item Over..."
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdCount 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Count Items..."
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdClear1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Clear Listbox..."
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete Item..."
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   840
      Width           =   2175
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Add Item..."
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   360
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
      TabIndex        =   13
      Top             =   6600
      Width           =   5295
   End
   Begin VB.Frame fraOther 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Other"
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   6015
      Begin VB.ListBox lstListAddFrom 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1230
         ItemData        =   "frmListbox.frx":0FB4
         Left            =   2400
         List            =   "frmListbox.frx":0FD3
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtListAddTo 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   4200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Textbox:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   4200
         TabIndex        =   17
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Listbox:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame fraBasics 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Basics"
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   6015
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1815
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.Frame fraFile 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save and Load"
      ForeColor       =   &H00000000&
      Height          =   2175
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   6015
      Begin VB.ListBox List2 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1815
         ItemData        =   "frmListbox.frx":1003
         Left            =   2400
         List            =   "frmListbox.frx":1016
         TabIndex        =   8
         Top             =   240
         Width           =   3495
      End
   End
End
Attribute VB_Name = "frmListbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
    
    Dim tempString As String
    
    tempString = InputBox("Input text for a new entry:", "Learn Visual Basic")
    List1.AddItem tempString
    
End Sub

Private Sub cmdAddOver_Click()
    
    Dim i As Integer
    
    For i = 0 To lstListAddFrom.ListCount
        txtListAddTo.Text = txtListAddTo.Text & lstListAddFrom.List(i) & vbNewLine
    Next
    
End Sub

Private Sub cmdClear1_Click()
    
    List1.Clear
    
End Sub

Private Sub cmdClear2_Click()
    
    List2.Clear
    
End Sub

Private Sub cmdCount_Click()
    
    MsgBox "There are " & List1.ListCount & " items.", vbInformation, "Learn Visual Basic"
    
End Sub

Private Sub cmdDelete_Click()
    
    On Error Resume Next
    
    List1.RemoveItem List1.ListIndex
    
End Sub

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdMoveOver_Click()
    
    txtListAddTo.Text = txtListAddTo.Text & lstListAddFrom.List(lstListAddFrom.ListIndex)
    
End Sub

Private Sub cmdSave_Click()
    
    Dim i As Integer
    
    Open App.Path & "\Dummy Files\Listbox Test.txt" For Output As #1
            
        For i = 0 To List2.ListCount - 1
            Print #1, (List2.List(i))
        Next
        
    Close #1
    
End Sub

Private Sub Command1_Click()
    
    Dim tempString As String
    
    Open App.Path & "\Dummy Files\Listbox Test.txt" For Input As #1
        
        While Not EOF(1)
            Input #1, tempString
            List2.AddItem tempString
        Wend
        
    Close #1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "Listbox"
    
End Sub
