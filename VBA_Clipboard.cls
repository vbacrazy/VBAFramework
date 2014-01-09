VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBA_Clipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' .NET Framework System.Windows.Forms.ClipboardÇÃÉNÉçÅ[Éì
'
' Author: D*isuke YAMAKAWA
' last modified : 2014/1/7
Private DOClipboard
Private Sub Class_Initialize()
    Set DOClipboard = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
End Sub

Function GetText()
    With DOClipboard
        .GetFromClipboard
        GetText = .GetText
    End With
End Function

Sub SetText(str As String)
    With DOClipboard
        .SetText str
        .PutInClipboard
    End With
End Sub

Sub Clear()
    With DOClipboard
        .Clear
    End With
End Sub

