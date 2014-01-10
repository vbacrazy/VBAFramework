VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBA_Regex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' System.Text.RegularExpressions.Regex ÇÃÉNÉçÅ[Éì
'
' Author: D*isuke YAMAKAWA
' last modified : 2014/1/10
Public Enum RegexOptions
    IgnoreCase = &H80000000
    MultiLine = &H80000001
End Enum

Private re As Object

Public Sub Init(pattern As String, Optional options As String)
    Set re = CreateObject("VBScript.RegExp")
    re.pattern = pattern
    re.IgnoreCase = IIf(InStr(options, RegexOptions.IgnoreCase), True, False)
    re.MultiLine = IIf(InStr(options, RegexOptions.MultiLine), True, False)
End Sub

Public Function IsMatch(input_ As String, Optional startAt As Integer = 1)
    IsMatch = re.test(Mid(input_, startAt))
End Function


Public Function Match(input_ As String, Optional startAt As Integer = 1)
    re.Global = False
    Set Match = re.Execute(Mid(input_, startAt))(0)
End Function

Public Function Matches(input_ As String, Optional startAt As Integer = 1)
    re.Global = True
    Set Matches = re.Execute(Mid(input_, startAt))
End Function

Public Function Replace(input_ As String, replacement As String)
    Replace = re.Replace(input_, replacement)
End Function


