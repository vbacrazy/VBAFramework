VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBA_VAMIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' VAMIE (VBA Auto Mation for Internet Explorer)
'
' LastModified:
'   2014/1/8    + DOMセレクタ拡張 (ByNameを追加)
'
'   2014/1/7    * Vbagetシリーズとしてリファイン(分散していたコードを集積)
'                   + プロパティの整備
'                   + 特殊コマンドの追加(goto_url_NoWait, Wait, DisableConfirmFunction)
'                   + 64bitへの対応
'
'   2012/6/20   - domselec()のエラー処理削除
'   2012/6/18   + domselec()にエラー処理追加 （デバッグ時に問題を把握しづらいため）
'
' This Class Module:
'   First Release 2011/11/08
'   Created By D*isuke YAMAKWA
'
' SpecialThanks:
'   Excel VBAのマクロで, IEを自動操作しよう(DOMセレクタ関数をVBAで自作)
'   http://d.hatena.ne.jp/language_and_engineering/20090710/p1
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private ie As Object

Property Let Visible(x As Boolean)
    ie.Visible = x
End Property
Property Get Visible() As Boolean
    Visible = ie.Visible
End Property

Property Get Document() As Boolean
    Document = ie.Document
End Property


Sub Class_Initialize()
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
End Sub

Sub goto_url(url)
    ie.Navigate url
    waitIE
End Sub

' waitIEを挟むと無限ループするようなページ対策
Sub goto_url_NoWait(url)
    ie.Navigate url
End Sub

Sub Quit()
    ie.Quit
End Sub

' $
Function gid(dom_id)
    ' 注：旧IEのgetElementByIdはnameも参照する
    Set gid = ie.Document.getElementById(dom_id)
End Function

' getElementsByTagName
Function gtn(parent, tag_name)
    Set gtn = parent.getElementsByTagName(tag_name)
End Function

' 入力します
Sub type_val(dom_id, val)
    gid(dom_id).value = val
    Sleep 100
End Sub

' 送信ボタンやリンクをクリック
Sub submit_click(dom_id)
    gid(dom_id).Click
    waitIE
End Sub

' 簡易DOMセレクタ
Function domselec(arr)
    Dim parent_obj As Object: Set parent_obj = ie.Document
    Dim child_obj As Object
    Dim cur, continue_flag, dom_id, tag_name, index_num, name_

    cur = 0: continue_flag = True
    
    Do While continue_flag = True
        Select Case arr(cur):
            Case "id"
                dom_id = arr(cur + 1)
                Set child_obj = parent_obj.getElementById(dom_id)
                cur = cur + 2
            Case "tag"
                tag_name = arr(cur + 1)
                index_num = arr(cur + 2)
                Set child_obj = parent_obj.getElementsByTagName(tag_name)(index_num)
                cur = cur + 3
            Case "name"
                name_ = arr(cur + 1)
                index_num = arr(cur + 2)
                Set child_obj = parent_obj.getElementsByName(name_)(index_num)
                cur = cur + 3
        End Select
        
        Set parent_obj = child_obj
        
        If cur > UBound(arr) Then
            continue_flag = False
        End If
    Loop
    
    Set domselec = parent_obj
End Function

' 要素をクリックします
Sub ie_click(dom_id)
    gid(dom_id).Click
    Sleep 100
End Sub


' チェックボックスの状態をセットします
Sub set_check_state(dom_id, checked_flag)
    ' 希望通りのチェック状態でなければクリック
    If Not (gid(dom_id).Checked = checked_flag) Then
        ie_click dom_id
    End If
End Sub


' セレクトボックスを文言ベースで選択します
Sub select_by_label(dom_id, label)
    If Len(label) < 1 Then
      Exit Sub
    End If
    
    Dim opts As Object
    Dim i As Integer
    Set opts = gid(dom_id).Options
    For i = 0 To opts.Length - 1
        ' textが同じか
        If opts(i).innerText = label Then
            opts(i).Selected = True
            Exit Sub
        End If
    Next i
    
End Sub

' ラジオボタンを値ベースで選択します
Sub select_radio_by_val(post_name, value)
    If Len(value) < 1 Then
        Exit Sub
    End If
    
    Dim radios: Set radios = ie.Document.getElementsByName(post_name)
    Dim i: For i = 0 To radios.Length - 1
        If radios(i).value = CStr(value) Then
            radios(i).Click
            Sleep 100
        End If
    Next i

End Sub
' =======================================================
Sub waitIE()
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop
    
    Sleep 100
End Sub

Sub Wait(millisecond As Integer)
    Sleep millisecond
    DoEvents
End Sub

' =======================================================
Function GetIEVer()
    Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
    Dim hoge: hoge = Fix(val(FS.GetFileVersion(ie.FullName)))
    GetIEVer = hoge
End Function

'confirm()呼び出し時に確認ダイアログを表示させない
Sub DisableConfirmFunction()
    Dim ele: Set ele = ie.Document.createElement("SCRIPT")
    
    ele.Type = "text/javascript"
    ele.text = "function confirm() { return true; }"
    
    Call ie.Document.body.appendChild(ele)
End Sub
