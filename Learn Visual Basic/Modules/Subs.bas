Attribute VB_Name = "Subs"
Option Explicit

Public Sub ShowHelp(API As String, NameOfFile As String)
    
    If API = "None" Then
        frmHelp.txtAPI.Text = "None"
    Else
        frmHelp.txtAPI.LoadFile App.Path & "\Files\" & API & " API.rtf"
    End If
    
    frmHelp.txtCode.LoadFile App.Path & "\Files\" & NameOfFile & " Help.rtf"
    
    frmHelp.txtCode.SelStart = 0
    frmHelp.txtCode.SelLength = Len(frmHelp.txtCode.Text)
    frmHelp.txtCode.SelColor = &H808080
    frmHelp.txtCode.SelStart = 0
    
    frmHelp.txtAPI.SelStart = 0
    frmHelp.txtAPI.SelLength = Len(frmHelp.txtAPI.Text)
    frmHelp.txtAPI.SelColor = &H808080
    frmHelp.txtAPI.SelStart = 0
    
    frmHelp.Show
    
End Sub
