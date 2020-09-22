VERSION 5.00
Begin VB.Form frmIfThen 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "If ... Then"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6255
   Icon            =   "frmIfThen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5295
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
      TabIndex        =   1
      Top             =   4560
      Width           =   5295
   End
   Begin VB.PictureBox picHelp 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   5595
      Picture         =   "frmIfThen.frx":08CA
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4635
      Width           =   495
   End
   Begin VB.TextBox txtVar1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   480
      MaxLength       =   2
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox txtVar2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "10"
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton cmdExecute 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Execute..."
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3960
      Width           =   2175
   End
   Begin VB.Frame fraCode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Code"
      ForeColor       =   &H00000000&
      Height          =   4335
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6015
      Begin VB.Label lblVar2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1620
         TabIndex        =   14
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblVar1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   13
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "End If"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   240
         TabIndex        =   12
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MsgBox ""Equal"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   720
         TabIndex        =   11
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Else"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   240
         TabIndex        =   10
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MsgBox ""Less Than"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   9
         Top             =   1800
         Width           =   3255
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "MsgBox ""Greater Than"""
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   720
         TabIndex        =   8
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ElseIf       <       Then"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label lblLine 
         BackColor       =   &H00FFFFFF&
         Caption         =   "If       >       Then"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmIfThen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDone_Click()
    
    Unload Me
    
End Sub

Private Sub cmdExecute_Click()
    
    If Not IsNumeric(txtVar1) Or Not IsNumeric(txtVar2) Then
        MsgBox "Invalid inputs.", vbExclamation, "Learn Visual Basic"
        Exit Sub
    End If
    
    If Val(txtVar1) > Val(txtVar2) Then
        MsgBox "Greater Than.", vbInformation, "Learn Visual Basic"
    ElseIf Val(txtVar1) < Val(txtVar2) Then
        MsgBox "Less Than.", vbInformation, "Learn Visual Basic"
    Else
        MsgBox "Equal.", vbInformation, "Learn Visual Basic"
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    frmMain.Show
    
End Sub

Private Sub picHelp_Click()
    
    ShowHelp "None", "If Then"
    
End Sub

Private Sub txtVar1_Change()
    
    lblVar1.Caption = txtVar1.Text
    
End Sub

Private Sub txtVar2_Change()
    
    lblVar2.Caption = txtVar2.Text
    
End Sub
