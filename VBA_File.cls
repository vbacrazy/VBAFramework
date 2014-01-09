VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBA_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' .NET Framework System.IO.Fileのクローン
' ※　ただし、ファイルの読み書きはShift-JISしか扱えない
'
' Author: D*isuke YAMAKAWA
' last modified : 2013/10/9

Public Sub AppendAllText(path As String, contents As String)
    On Error GoTo SkipW
    Dim fileNum As Integer: fileNum = FreeFile
    
    Open path For Append As #fileNum
    Print #fileNum, contents
SkipW:
    Close #fileNum
End Sub

Public Sub Copy(sourceFileName As String, destFileName As String, Optional overwrite As Boolean = False)
    If Exists(destFileName) Then
        If overwrite Then
            Call FileCopy(sourceFileName, destFileName)
        End If
    Else
        Call FileCopy(sourceFileName, destFileName)
    End If
End Sub

Public Sub Delete(path As String)
    Call Kill(path)
End Sub

Public Function Exists(filePath As String)
    With CreateObject("Scripting.FileSystemObject")
        If .FileExists(filePath) Then
            Exists = True
        Else
            Exists = False
        End If
    End With
End Function

Public Sub Move(sourceFileName As String, destFileName As String)
    With CreateObject("Scripting.FileSystemObject")
        Call .MoveFile(sourceFileName, destFileName)
    End With
End Sub

'----------------------------------------------------------------
' ファイルを読み込む　（内容を文字列にして返す）
'----------------------------------------------------------------
Public Function ReadAllText(fileName As String)
    Dim fso As Object, buf As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    With fso.GetFile(fileName).OpenAsTextStream
        buf = .ReadAll
        .Close
    End With
    Set fso = Nothing
    
    ReadAllText = buf
End Function

' 末尾に変な空白文字が入る
Public Function ReadAllLines(fileName As String)
    Dim fso As Object, buf As String
    Dim data() As String
    ReDim data(0)
    Dim i As Integer
    i = 0
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    With fso.GetFile(fileName).OpenAsTextStream
        Do Until .AtEndOfStream
            data(i) = .ReadLine
            i = i + 1
            ReDim Preserve data(i)
        Loop
        
        .Close
    End With
    Set fso = Nothing
    
    ReadAllLines = data
End Function

Public Sub WriteAllLines(path As String, contents() As String)
    On Error GoTo SkipW
    Dim fileNum As Integer: fileNum = FreeFile
    
    Open path For Output As #fileNum
    Dim c: For Each c In contents
        Print #fileNum, c
    Next
SkipW:
    Close #fileNum
End Sub


Public Sub WriteAllText(path As String, contents As String)
    On Error GoTo SkipW
    Dim fileNum As Integer: fileNum = FreeFile
    
    Open path For Output As #fileNum
    Print #fileNum, contents
SkipW:
    Close #fileNum
End Sub
