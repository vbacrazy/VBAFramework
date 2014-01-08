VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBA_Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' .NET Framework System.Diagnostics.Stopwatch ÇÃÉNÉçÅ[Éì
'
' Author: D*isuke YAMAKAWA
' last modified : 2014/1/8

#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As LongPtr
#Else
    Private Declare Function GetTickCount Lib "kernel32" () As Long
#End If

Private StartTime As Date
Private EndTime As Date
Private AccumulatedTime As Date

Private Sub Class_Initialize()
End Sub

' --------------------------------------------------------
Property Get Elapsed()
    Elapsed = TimeSerial( _
                    AccumulatedTime / 1000 / 60 / 60, _
                    AccumulatedTime / 1000 / 60, _
                    AccumulatedTime / 1000)
End Property

Property Get ElapsedMilliseconds()
    ElapsedMilliseconds = AccumulatedTime
End Property
' --------------------------------------------------------

Public Sub Start()
    StartTime = GetTickCount()
End Sub

Public Sub StartNew()
    StartTime = GetTickCount()
    AccumulatedTime = 0
End Sub

Public Sub Stop_()
    EndTime = GetTickCount()
    
    AccumulatedTime = AccumulatedTime + EndTime - StartTime
End Sub

Public Sub Reset()
    AccumulatedTime = 0
End Sub

Public Sub Restart()
    Call StartNew
End Sub
