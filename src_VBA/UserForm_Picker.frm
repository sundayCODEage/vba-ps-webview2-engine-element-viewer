VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Picker 
   Caption         =   "画面エクスポート ＆ ピッカー"
   ClientHeight    =   6105
   ClientLeft      =   108
   ClientTop       =   5400
   ClientWidth     =   6480
   OleObjectBlob   =   "UserForm_Picker.frx":0000
End
Attribute VB_Name = "UserForm_Picker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==============================================================================
' [API宣言] 最前面表示・スリープ処理用
Private Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hwnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal uFlags As Long) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' 連想配列(今回は、「Microsoft Scripting Runtime」利用していない）
Private dicUrls As Object
' ==============================================================================

' --- [イベント] フォーム起動時 ---
Private Sub UserForm_Initialize()
    ' フォームの表示位置:「Manual に設定」
    Me.StartUpPosition = 0
    ' 「右下」にずらす
    Me.Left = Application.Left + Application.Width - Me.Width - 60
    Me.Top = Application.Top + Application.Height - Me.Height - 60

    ' Excel本体の最小化
    Application.WindowState = xlMinimized
    ' デフォルトで最前面表示を有効にする
    Me.chkTopMost.Value = True
    Call SetTopMost(True)
    
    ' --- 選択の待機時間コンボボックスの初期化 ---
    With Me.cmbWaitSec
        .Clear
        .AddItem "3"
        .AddItem "5"
        .AddItem "8"
        .Value = "3"    ' デフォルトを3秒に設定する
    End With
    
    ' --- 要素の待機時間コンボボックスの初期化 ---
    With Me.cmbTimeOut
        .Clear
        .AddItem "30"
        .AddItem "45"
        .AddItem "60"
        .Value = "45"   ' デフォルトを30秒に設定する
    End With
    
    ' --- 連想配列にテスト対象を登録する。 ---  (Key: コンボボックスの表示名, Item: 実際のURL)
    Set dicUrls = CreateObject("Scripting.Dictionary")
    
    Dim ws As Worksheet
    Dim r As Long
    Dim testId As String
    Dim url As String
    Dim testNote As String
    Dim dispName As String
    
    ' テストケースのシートとデータ開始行を指定
    Set ws = ThisWorkbook.Sheets("TestCases")
    r = glb_TestLoop_STA
    Me.cmbUrlList.Clear
    
    ' A列(TestID)が空白になるまでループ
    Do While ws.Cells(r, 1).Value <> ""
        testId = ws.Cells(r, 1).Value
        url = ws.Cells(r, 2).Value
        testNote = ws.Cells(r, 6).Value
        ' URLが入力されていればリストに登録する
        If url <> "" Then
            ' コンボボックスの表示名を作成
            dispName = testId & " : " & testNote
            ' 辞書のキー重複エラーを防止
            If Not dicUrls.Exists(dispName) Then
                dicUrls.Add dispName, url
                Me.cmbUrlList.AddItem dispName
            End If
        End If
        r = r + 1
    Loop

End Sub

' --- [イベント] フォーム終了時 ---
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' フォームを閉じたら、Excel本体を元のサイズ（最大化）に戻す
    Application.WindowState = xlMaximized
    ' [メインループ]へ: 終了を通知
    glb_DebugAction = 9
End Sub

' --- [イベント] 最前面表示チェックボックスの切り替え ---
Private Sub chkTopMost_Click()
    Call SetTopMost(Me.chkTopMost.Value)
End Sub

' --- 最前面／解除を切り替える ---
Private Sub SetTopMost(ByVal isTop As Boolean)
    Dim hwnd As LongPtr
    hwnd = FindWindow("ThunderDFrame", Me.Caption)

    If hwnd <> 0 Then
        If isTop Then
            ' 最前面にピン留め
            Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        Else
            ' ピン留め解除（裏に回れる）
            Call SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        End If
    End If
End Sub

' --- [イベント] エクスポート実行ボタン ---
Private Sub btnExport_Click()
    MsgBox "エクスポートを実行します", vbInformation
    
    Debug.Print "=========================================================="
    Debug.Print " >> btnExport_Click時の、Tabs一覧を表示"
    Call btnGetTabs_Click
    
    ' [メインループ]へ: エクスポート実行を依頼
    glb_DebugAction = 1
    ' エクスポート対象のチェック状態を退避
    glb_CheckBox1 = Me.CheckBox1.Value
    glb_CheckBox2 = Me.CheckBox2.Value
    glb_CheckBox3 = Me.CheckBox3.Value
    glb_CheckBox4 = Me.CheckBox4.Value
    glb_CheckBox5 = Me.CheckBox5.Value
    ' エクスポートする画面内容のメモを退避
    glb_DebugNote = Me.TextNote.Text
    
    Me.lblStatus.Caption = "出力処理中..."
    DoEvents
End Sub

' --- タブ一覧の取得とリストボックスへの表示 ---
Private Sub btnGetTabs_Click()
    Dim resJson As String
    Dim jsonParsed As Object
    Dim item As Object

    Me.lblStatus.Caption = "タブ一覧を取得中..."
    DoEvents
    
    ' RPAエンジンからタブ一覧を取得
    resJson = GetRpaEngine().RunAction("List-Tabs", CreateParams())
    
    ' リストボックスを初期化 (2列設定: 0列目=Title, 1列目=TabId)
    lstTabs.Clear
    lstTabs.ColumnCount = 2
    lstTabs.ColumnWidths = "150;0" ' 2列目は非表示

    ' JSONをパースしてリストに追加
    Set jsonParsed = JsonConverter.ParseJson(resJson)
    If Not jsonParsed Is Nothing Then
        For Each item In jsonParsed
            lstTabs.AddItem item("Title")
            lstTabs.List(lstTabs.ListCount - 1, 1) = item("TabId")
        Next item
        Me.lblStatus.Caption = "タブ取得完了: " & lstTabs.ListCount & " 件"
    End If
End Sub

' --- [イベント] ピッカー実行ボタン ---
Private Sub btnUiaPicker_Click()
    ' 連打防止のためボタンを一時的に無効化（メインループ側で復帰）
    Me.btnUiaPicker.Enabled = False
    ' 選択ハイライト
    glb_UseHighlight = Me.chkHighlight.Value
    ' 座標逆引きモード
    glb_UsePointReverse = Me.chkPointReverse.Value
    ' ピッカー選択の待機時間を数値に変換して退避
    glb_WaitSec = CInt(Me.cmbWaitSec.Value)
    ' 要素の取得係る待機時間を数値に変換して退避
    glb_TimeOut = CInt(Me.cmbTimeOut.Value)

    ' [メインループ]へ: ピッカー実行を依頼
    glb_DebugAction = 2
End Sub

' --- [選択Ur]コンボボックスでの項目選択 ---
Private Sub cmbUrlList_Change()
    Dim selectedUrl As String
    Dim sandboxPath As String
    ' 対応するURLをテキストボックスにセットする
    If Not dicUrls Is Nothing Then
        If dicUrls.Exists(Me.cmbUrlList.Value) Then
            selectedUrl = dicUrls(Me.cmbUrlList.Value)
            ' ローカルパスの置換処理
            If InStr(selectedUrl, "ローカルパス") > 0 Then
                sandboxPath = Replace(ThisWorkbook.Path & "\sandbox\", "\", "/")
                selectedUrl = "file:///" & sandboxPath & Replace(selectedUrl, "[ローカルパス]/", "")
            End If

            Me.TextUrl.Text = selectedUrl
        End If
    End If
End Sub

' --- [選択Ur]クリップボード経由 ---
Private Sub TextUrl_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    ' クリップボードのテキストを直接貼り付け
    Me.TextUrl.Paste
End Sub

' --- [選択Ur]選択したUrlへ移動 ---
Private Sub btnNavigate_Click()
    ' URLが空でなければ移動フラグ(8)を立てる
    If Trim(Me.TextUrl.Text) <> "" Then
        glb_DebugAction = 8
    Else
        MsgBox "URLを入力してください。", vbExclamation
    End If
End Sub

' --- [イベント] 終了ボタン ---
Private Sub btnClose_Click()
    ' [メインループ]へ: 終了を通知
    glb_DebugAction = 9
    Unload Me
End Sub

' ==============================================================================
' --- テストボタン ---
Private Sub TEST_Button_Click()
    ' xxx
End Sub
