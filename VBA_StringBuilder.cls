VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBA_StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' .NET Framework System.Text.StringBuilderÇÃÉNÉçÅ[Éì
'
' Author: D*isuke YAMAKAWA
' last modified : 2014/1/9

Private mBuffer As Object
Private mIndex As Long

Private Sub Class_Initialize()
    Set mBuffer = CreateObject("Scripting.Dictionary")
    mIndex = 0
End Sub
 
Private Sub Class_Terminate()
    Set mBuffer = Nothing
End Sub
 
Public Sub Append(value As String)
    mIndex = mIndex + 1
    Call mBuffer.Add(CStr(mIndex), value)
End Sub
 
Public Sub AppendLine(Optional value As String = "")
    mIndex = mIndex + 1
    Call mBuffer.Add(CStr(mIndex), value & vbCrLf)
End Sub
 
Public Sub Clear()
    mIndex = 0
    mBuffer.RemoveAll
End Sub

Public Function ToString() As String
    ToString = Join(mBuffer.Items(), "")
End Function
