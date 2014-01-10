VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBA_ConfigFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' ä»à’ConfigFileì«Ç›èëÇ´ÉNÉâÉX
'
' Author: D*isuke YAMAKAWA
' last modified : 2014/1/10

Private FileName_ As String
Property Get FileName() As String
    FileName = FileName_
End Property
Property Get FilePath() As String
    Dim Wsh: Set Wsh = CreateObject("WScript.Shell")
    Dim path As String: path = Wsh.SpecialFolders("MyDocuments") & "\MyMacroConfig"
    
    If Dir(path, vbDirectory) = "" Then MkDir (path)
    
    FilePath = path & "\" & FileName
End Property

Public Sub Init(fileNameOnly As String)
    FileName_ = fileNameOnly
End Sub

' Shift-JISÇ≈èëÇ´çûÇ›
Public Sub Save(data() As String)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim stream: Set stream = fso.OpenTextFile(ConfigFilePath, 2, True, 0)
    
    Dim data_item: For Each data_item In data
        Call stream.WriteLine(data_item)
    Next
    
    stream.Close
End Sub

' Shift-JISÇ≈ì«Ç›çûÇ›
Public Function Load() As Variant
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Dim stream: Set stream = fso.OpenTextFile(ConfigFilePath, 1, True, 0)
    ReadConfigFile = Split(stream.ReadAll, vbCrLf)
    stream.Close
End Function

