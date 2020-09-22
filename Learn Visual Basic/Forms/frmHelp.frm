VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   8295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraCode 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Code Used"
      ForeColor       =   &H00000000&
      Height          =   5535
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   8055
      Begin RichTextLib.RichTextBox txtCode 
         Height          =   5145
         Left            =   135
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   255
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   9075
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmHelp.frx":08CA
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H80000003&
         Height          =   5175
         Left            =   120
         Top             =   240
         Width           =   7815
      End
   End
   Begin VB.Frame fraAPI 
      BackColor       =   &H00FFFFFF&
      Caption         =   "API / Component Add-in"
      ForeColor       =   &H00000000&
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin RichTextLib.RichTextBox txtAPI 
         Height          =   1425
         Left            =   135
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   255
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   2514
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmHelp.frx":094C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H80000003&
         Height          =   1455
         Left            =   120
         Top             =   240
         Width           =   7815
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
