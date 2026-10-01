Attribute VB_Name = "Mod_TestDevTools"
Option Explicit

' ==============================================================================
' [API宣言] 高精度タイマー・スリープ処理用
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ==============================================================================
' [モジュールレベル変数]
Private freq As Currency            ' CPUの動作周波数 (タイマー精度)
    Dim startTime As Currency
    Dim endTime As Currency
    Dim elapsedTimeMs As Double

' [◆] 他モジュールからの干渉を防ぐため Private を維持する
Private rpaEngine As Ps_Engine      ' RPA操作エンジン本体クラス

' ==============================================================================
' エンジン設定
Private Const PS_ENGINE_LIB As String = "\Ps_Engine_Core_v301.ps1"

' ==============================================================================
    Dim loopCount As Integer        ' デバッグ用ループカウンタ
Public glb_WaitSec As Integer       ' ピッカー選択の待機時間
Public glb_TimeOut As Integer       ' 要素の取得係る待機時間
Public Const glb_TestLoop_STA As Integer = 11   ' 自動テスト（開始R）

' [グローバル変数] ユーザーフォームと通信するためのフラグ・データ
Public glb_DebugAction As Integer       ' (0:待機中, 1:エクスポート, 2:ピッカー, 8:URL遷移, 9:終了)

Public glb_UseHighlight     As Boolean  ' 選択時ハイライト
Public glb_UsePointReverse  As Boolean  ' 座標逆引きモード

Public glb_DebugNote As String          ' エクスポート時の付記メモ
Public glb_CheckBox1 As Boolean         ' エクスポート対象: Html
Public glb_CheckBox2 As Boolean         ' エクスポート対象: Screens
Public glb_CheckBox3 As Boolean         ' エクスポート対象: Elements
Public glb_CheckBox4 As Boolean         ' エクスポート対象: DomSnapshot
Public glb_CheckBox5 As Boolean         ' エクスポート対象: Layout

' ==============================================================================
' [Getter] フォームからPrivate変数(rpaEngine)へアクセスする
Public Function GetRpaEngine() As Ps_Engine
    Set GetRpaEngine = rpaEngine
End Function

' ==============================================================================
'   DevTools／テスト (画面情報エクスポート＆要素ピッカー)
' ==============================================================================
Sub Test_DevTools()

    Dim enginePath As String
    Dim sessionId As String
    Dim useCdpPort As Integer   ' 通信モードの設定 (0: 標準, 9222: CDP高速通信)
    
    ' CPUのタイマー周波数を取得 (高精度計測の準備)
    QueryPerformanceFrequency freq
    
    ' --- 実行環境・セッションの初期化 ---
    enginePath = ThisWorkbook.Path & PS_ENGINE_LIB
    sessionId = "SESSION_" & Format(Now, "yyyyMMdd_HHmmss")
    ' --- CDP通信モードの設定 ---
    ' 通常の運用では標準モード(0)を推奨しますが、本テストシナリオでは
    ' CDP機能(Invoke-CdpScript等)を検証するため、ポート番号(例: 9222)を指定します。
    Debug.Print ".>> 通常の運用では標準モード(0)を推奨、テストではポート番号を指定します。"
    useCdpPort = 9222   ' CDP高速通信ポート (0指定で標準モード)
    
    Set rpaEngine = New Ps_Engine

    ' --- PowerShellエンジンを起動 ---
    ' 第3引数: CDPポート番号
    ' 第4引数: IsDebugModeFlg (True = 実行時のパラメータや詳細ログをSTDOUTへ出力する)
    If Not rpaEngine.StartEngine(sessionId, enginePath, useCdpPort, True) Then
        MsgBox "RPAエンジンの起動に失敗しました。", vbCritical, "起動エラー"
        Exit Sub
    End If
    
    ' ==========================================================================
    ' --- テスト画面の呼び出し ---
    ' ==========================================================================
    ' Excelを強制的に最前面に持ってくる
    Call ForceFocusExcel

    Dim prompt As String
    Dim ans As String

    prompt = "パート（テスト画面）を選んでください。" & vbCrLf & _
            " 11. sandbox:(01_basic_form)画面" & vbCrLf & _
            " 12. sandbox:(12_scroll_shadow_dom)画面" & vbCrLf & _
            " ---" & vbCrLf & _
            " 21. ウィキペディアへようこそ" & vbCrLf & _
            " 22. Yahoo!路線情報" & vbCrLf & _
            " 23. youtube.com" & vbCrLf & _
            " ---" & vbCrLf & _
            " 31. DevTools自動テスト" & vbCrLf & _
            " ---" & vbCrLf & _
            " 88. Url選択(TestCasesシートのUrlほか）" & vbCrLf & _
            " 99. xxx( TEST )"
             
    ans = InputBox(prompt, "テストパート選択")
    
    If ans = "" Then
        MsgBox "キャンセルされました。", vbInformation
        GoTo CleanUp
    End If
    
    MsgBox "選択したテストパート: ( " & ans & " )", vbInformation
    If ActiveSheet.Name = "TestCases" Then
        If ans = "31" Then
            MsgBox "「TestCases」は、事前に PC環境に併せた座標登録(x,y)が必要です。", vbInformation
        Else
            MsgBox "保護エラー：「TestCases」シートがアクティブな状態では起動できません。" & vbCrLf & _
                    "（ピッカーの結果で自動テストデータが上書きされるのを防ぐ）", vbExclamation
            GoTo CleanUp
        End If
    End If

    ' 裏に隠れてしまったブラウザ（RPA Browser）を一番手前に呼び戻す
    On Error Resume Next
    AppActivate "RPA Browser"
    On Error GoTo 0
    ' 画面の切り替えが完全に落ち着くまでn秒待つ
    Application.Wait Now + TimeValue("00:00:01")

    ' 実行時のエンジン動作設定 (要素ハイライト機能の制御: True=有効)
    ' .. (有効関数: Invoke-WebXPathClick、Set-WebXPathTextInput、WebXPathText)
    rpaEngine.RunAction "Set-EngineConfig", CreateParams("EnableHighlight", True)
    
    Dim targetUrl As String
    Dim sandboxPath As String
    ' サンドボックス（ローカルHTML）の絶対パスを取得 (Windowsの \ を / に変換)
    sandboxPath = Replace(ThisWorkbook.Path & "\sandbox\", "\", "/")
    
    ' ==========================================================================
    ' --- テスト実行 ---
    Select Case ans
        Case "11"
            targetUrl = "file:///" & sandboxPath & "01_basic_form.html"
        Case "12"
            targetUrl = "file:///" & sandboxPath & "12_scroll_shadow_dom.html"
        Case "21"
            targetUrl = "https://ja.wikipedia.org/wiki/"
        Case "22"
            targetUrl = "https://transit.yahoo.co.jp/"
        Case "23"
            targetUrl = "https://www.youtube.com/"
        Case "31"
            Call Test_DevTools_AutoTests
            GoTo CleanUp
        Case "88"
            targetUrl = "about:blank"
            Call Test_DevTools_88(targetUrl)
        Case "99" ' テスト用ダミー
'            targetUrl = "https://**"
            targetUrl = "http://win-926k0drcr26/MCNS01/mc/a1/login.jsp"
        Case Else
            MsgBox "無効な入力です。", vbExclamation
            GoTo CleanUp
    End Select

    On Error GoTo ErrorHandler
    ' ==========================================================================
    rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
    rpaEngine.RunAction "Wait-WebDocumentReady", CreateParams("TimeoutSec", 10)
    Sleep 500

    ' --- モードレス（非同期）でフォームを表示 ---
    glb_DebugAction = 0 ' 初期化
    UserForm_Picker.Show vbModeless
    loopCount = 1
    
    ' ==========================================================================
    ' --- 非同期監視ループ ---
    Do
        ' [● 1: エクスポート実行]
        If glb_DebugAction = 1 Then
            Debug.Print "--- TEST_Debughelper / loopCount: " & Format(loopCount, "00") & " Memo: " & glb_DebugNote
            glb_DebugAction = 0 ' フラグをリセット
            Call Debug_ExportLoop(loopCount, useCdpPort)
            UserForm_Picker.lblStatus.Caption = "T" & Format(loopCount, "00") & " 出力完了。"
            loopCount = loopCount + 1
        End If
        
        ' [● 2: ピッカー実行]
        If glb_DebugAction = 2 Then
            glb_DebugAction = 0 ' フラグをリセット
            Call Debug_PickerLoop
        End If
        
        ' [● 8: 任意のURLへナビゲーション]
        If glb_DebugAction = 8 Then
            glb_DebugAction = 0 ' フラグをリセット
            targetUrl = UserForm_Picker.TextUrl.Text ' フォームに入力された最新のURLを取得
            
            UserForm_Picker.lblStatus.Caption = "ページ移動中: " & targetUrl
            DoEvents
            
            rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", targetUrl)
            rpaEngine.RunAction "Wait-WebDocumentReady", CreateParams("TimeoutSec", 10)
            Sleep 500
            
            ' 裏に隠れてしまったブラウザ（RPA Browser）を一番手前に呼び戻す
            On Error Resume Next
            AppActivate "RPA Browser"
            On Error GoTo ErrorHandler ' エラー処理へ

            UserForm_Picker.lblStatus.Caption = "移動完了。ピッカーを実行できます。"
        End If
        
        ' [● 9: デバッグ終了]
        If glb_DebugAction = 9 Then
            Exit Do
        End If
        
        DoEvents
        Sleep 50
    Loop

    On Error GoTo ErrorHandler
    ' ==========================================================================
    ' --- テスト完了処理 ---
    Debug.Print "=== テスト完了 ==="
    MsgBox "テストが完了しました。", vbInformation

CleanUp:
    If Not rpaEngine Is Nothing Then
        rpaEngine.CloseEngine
    End If
    Set rpaEngine = Nothing
    Application.StatusBar = False
    Exit Sub
    
' ------------------------------------------------------------------
' [例外処理] 異常系のハンドリング
' ------------------------------------------------------------------
ErrorHandler:
    If Err.Source = "PS_Engine" Then
        Dim rpaErr As RpaExceptionInfo
        rpaErr = ParseRpaError(Err.Description)
        
        Debug.Print "【実行中エラー】 関数: " & rpaErr.FunctionName & " / 種別: " & rpaErr.ErrorType
        Debug.Print " メッセージ: " & rpaErr.Message
        Debug.Print " 詳細(Details): " & rpaErr.Details
        
        MsgBox "RPAエンジン実行中にエラーが発生しました。" & vbCrLf & _
               "関数: " & rpaErr.FunctionName & vbCrLf & _
               "内容: " & rpaErr.Message, vbCritical, "RPAシステムエラー"
    Else
        MsgBox "VBAマクロエラー (" & Err.Number & "): " & Err.Description, vbCritical
    End If
    
    On Error Resume Next
    rpaEngine.RunAction "Export-WebScreenshot", CreateParams("Prefix", "Test_ErrorDump")
    Resume CleanUp
End Sub

' ==============================================================================
' --- 画面情報エクスポート ---
' ==============================================================================
Private Sub Debug_ExportLoop(loopCount As Integer, useCdpPort As Integer)
    Dim fileName As String
    On Error Resume Next ' 取得できる要素は強制取得する

    If useCdpPort = 0 Then
        If glb_CheckBox4 = True Or _
            glb_CheckBox5 = True Then
            MsgBox "WebDomSnapshot (WebLayoutDump) は、CDP高速通信ポートで起動してください。", vbInformation
        End If
    End If
    ' --- Export-WebHtml ---
    If glb_CheckBox1 = True Then
        fileName = "T" & Format(loopCount, "00") & "-" & "WebHtml"
        rpaEngine.RunAction "Export-WebHtml", CreateParams("Prefix", fileName)
    End If
    ' --- Export-WebScreenshot ---
    If glb_CheckBox2 = True Then
        fileName = "T" & Format(loopCount, "00") & "-" & "WebScreenshot"
        rpaEngine.RunAction "Export-WebScreenshot", CreateParams("Prefix", fileName)
    End If
    ' --- Export-WebElementsToCsv ---
    If glb_CheckBox3 = True Then
        fileName = "T" & Format(loopCount, "00") & "-" & "WebElements" & ".csv"
        rpaEngine.RunAction "Export-WebElementsToCsv", CreateParams("FileName", fileName)
    End If
    ' --- Export-WebDomSnapshot ---
    If glb_CheckBox4 = True And useCdpPort > 0 Then
        ' CPUのタイマー周波数を取得 (高精度計測の準備)
        QueryPerformanceFrequency freq
    
        fileName = "T" & Format(loopCount, "00") & "-" & "DomSnapshot"
        rpaEngine.RunAction "Export-WebDomSnapshot", CreateParams("Prefix", fileName)
        QueryPerformanceCounter endTime
        elapsedTimeMs = (endTime - startTime) / freq * 1000
        Debug.Print " ◆ 計測/Export-WebDomSnapshot (TimeMs) : " & Format(elapsedTimeMs, "0.00")
    End If
    ' --- Export-WebLayoutDump ---
    If glb_CheckBox5 = True And useCdpPort > 0 Then
        fileName = "T" & Format(loopCount, "00") & "-" & "LayoutDump"
        rpaEngine.RunAction "Export-WebLayoutDump", CreateParams("Prefix", fileName)
    End If

    On Error GoTo 0
End Sub

' ==============================================================================
' --- 要素ピッカー（待機カウントダウン＆コード生成） ---
' ==============================================================================
Private Sub Debug_PickerLoop()
    Dim resultCode As String
    Dim i As Integer

    If rpaEngine Is Nothing Then
        MsgBox "エンジンが起動していません。", vbExclamation
        UserForm_Picker.btnUiaPicker.Enabled = True
        Exit Sub
    End If
    
    ' ==========================================================================
    ' クラスモジュール： Loop待機のタイムアウトを指定する。
    Debug.Print " ◆ Loop待機のタイムアウト: " & glb_TimeOut
    rpaEngine.Loop_TimeoutSec = glb_TimeOut

    ' ==========================================================================
    ' --- リアルタイム・ハイライト --- (黄色系は、視認性低い)
    Dim jsHighlightStart As String
    If glb_UseHighlight = True Then
        jsHighlightStart = "window._rpaOutline = document.createElement('div');" & _
        "window._rpaOutline.style.cssText = 'position:fixed; pointer-events:none; z-index:999999; outline:1px solid red; background:rgba(255,0,0,0.1); transition:all 0.05s;';" & _
        "document.body.appendChild(window._rpaOutline);" & _
        "window._rpaHover = function(e){ var el = document.elementFromPoint(e.clientX, e.clientY); if(el){ var r = el.getBoundingClientRect(); window._rpaOutline.style.top = r.top + 'px'; window._rpaOutline.style.left = r.left + 'px'; window._rpaOutline.style.width = r.width + 'px'; window._rpaOutline.style.height = r.height + 'px'; }};" & _
        "document.addEventListener('mousemove', window._rpaHover);"
        
        rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", jsHighlightStart)
    End If

    UserForm_Picker.txtResult.Text = ""
    ' ｎ秒カウントダウン（待機）
    For i = glb_WaitSec To 1 Step -1
        UserForm_Picker.lblStatus.Caption = i & " 秒前：対象の上にマウスを移動して待機..."
        DoEvents
        Sleep 1000 ' 1秒待機
    Next i
    
    ' --- 要素取得直前にハイライトイベントを解除・削除 ---
    Dim jsHighlightEnd As String
    If glb_UseHighlight = True Then
        jsHighlightEnd = "document.removeEventListener('mousemove', window._rpaHover); if(window._rpaOutline){ window._rpaOutline.remove(); }"
        rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", jsHighlightEnd)
    End If
    UserForm_Picker.lblStatus.Caption = "マウス下の要素を取得・解析中..."
    DoEvents
    
    ' ==========================================================================
    On Error Resume Next
    
    QueryPerformanceCounter startTime
    ' PowerShell側を待機無しへ(WaitSec:0)
    resultCode = rpaEngine.RunAction("Invoke-UiaRecordStep", CreateParams("WaitSec", 0, "UsePointReverse", glb_UsePointReverse))
    If Err.Number <> 0 Then
        Debug.Print "【ピッカー取得エラー】 " & Err.Description
        UserForm_Picker.txtResult.Text = "エラーが発生しました: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    QueryPerformanceCounter endTime
    elapsedTimeMs = (endTime - startTime) / freq * 1000
    Debug.Print " ◆ 計測/Invoke-UiaRecordStep_ (TimeMs) : " & Format(elapsedTimeMs, "0.00")
    ' 取得・生成されたコード
    UserForm_Picker.txtResult.Text = resultCode
    ' ==========================================================================
''' Debug.Print "resultCode  : " & resultCode
    If InStr(resultCode, "[RESULT]") > 0 Then
        ' 取得した要素情報をシートへ転記し、パースされた純粋なセレクタを受け取る
        Dim pureSelector As String
        pureSelector = WriteMappingToExcel(resultCode)
''' Debug.Print "pureSelector: " & pureSelector
    ' ==========================================================================
        ' --- セレクタ安定性検証 (テスト) ---
    ' 取得したセレクタが、動的ID(ext-gen-* 等)や仮想DOMの再描画によって
    ' 短時間で消滅・変化しないか(一過性の幻ではないか)を複数回検証し、
    ' RPA本番実行時の「要素が見つからない」エラー(Flakyテスト)を未然に防ぐ。
    ' ==========================================================================
        If pureSelector <> "" Then
            UserForm_Picker.lblStatus.Caption = "セレクタの安定性をテスト中..."
            DoEvents ' 画面描画を更新
            
            Dim stabilityRes As String
            On Error Resume Next
            ' 抽出した pureSelector を渡す
            stabilityRes = rpaEngine.RunAction("Test-SelectorStability", CreateParams("Selector", pureSelector, "Repeat", 3))
            
            If Err.Number <> 0 Then
                Debug.Print "【安定性テストNG】 " & Err.Description
                UserForm_Picker.lblStatus.Caption = "テストNG: 要素が不安定、または消失しました。"
                Err.Clear
            Else
                Debug.Print "【安定性テストOK】 " & stabilityRes
                UserForm_Picker.lblStatus.Caption = "コード生成・安定性テスト完了 (OK)"
            End If
            On Error GoTo 0
        Else
            UserForm_Picker.lblStatus.Caption = "コード生成完了 (セレクタ抽出スキップ)"
        End If
    Else
        UserForm_Picker.lblStatus.Caption = "要素の取得に失敗しました。"
    End If

    ' ボタンを再び有効化
    UserForm_Picker.btnUiaPicker.Enabled = True
End Sub

' --- 取得した要素情報を（ActiveSheet）に書き込む ---
Private Function WriteMappingToExcel(ByVal resultCode As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rawData As String
    Dim parsed As Object
    Dim clientX As Long, clientY As Long

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.rows.Count, 3).End(xlUp).Row + 1
    ' [RESULT] を取り除くと、残りは完全なJSON形式になる
    rawData = Trim(Replace(resultCode, "[RESULT]", ""))

    On Error GoTo ParseError
    
    ' 1. JSONとして直接パース
    Set parsed = JsonConverter.ParseJson(rawData)
    
    ' BoundingBoxからPC物理座標(ClientX/Y)を算出
    If TypeName(parsed("boundingBox")) = "Dictionary" Then
        Dim bbox As Object
        Set bbox = parsed("boundingBox")
        
        Dim dpr As Double: dpr = 1
        On Error Resume Next
        Dim dprRes As String
        dprRes = rpaEngine.RunAction("Invoke-WebScript", CreateParams("Js", "return window.devicePixelRatio || 1;"))
        dprRes = Trim(Replace(dprRes, "[RESULT]", ""))
        If IsNumeric(dprRes) Then dpr = CDbl(dprRes)
        On Error GoTo ParseError
        
        clientX = CLng((bbox("x") + (bbox("width") / 2)) * dpr)
        clientY = CLng((bbox("y") + (bbox("height") / 2)) * dpr)
    End If

    ' 2. Excel への書き込み処理
    ws.Cells(lastRow, 1).Value = parsed("selector")
    ws.Cells(lastRow, 2).Value = parsed("value")
    ws.Cells(lastRow, 3).Value = parsed("command")
    ws.Cells(lastRow, 4).Value = parsed("outerHtml")
    
    If Not bbox Is Nothing Then
        ws.Cells(lastRow, 5).Value = bbox("x")
        ws.Cells(lastRow, 6).Value = bbox("y")
        ws.Cells(lastRow, 7).Value = bbox("width")
        ws.Cells(lastRow, 8).Value = bbox("height")
    End If
    
    If clientX > 0 Or clientY > 0 Then
        ws.Cells(lastRow, 9).Value = clientX
        ws.Cells(lastRow, 10).Value = clientY
    End If

    ' テスト用に純粋なセレクタを戻り値として返す
    WriteMappingToExcel = parsed("selector")

    On Error GoTo 0
    Exit Function

ParseError:
    MsgBox "解析に失敗しました。生データ: " & vbCrLf & rawData, vbExclamation
    On Error GoTo 0
End Function

' ==============================================================================
' --- 取得座標の自動テスト ---: （シートに座標登録、シナリオの調整が必要）
' ==============================================================================
Private Sub Test_DevTools_AutoTests()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("TestCases")
    ' 結果列をクリア
    ws.Range("G11:H100").ClearContents
    
    Dim targetUrl As String
    Dim sandboxPath As String
    ' サンドボックス（ローカルHTML）の絶対パスを取得 (Windowsの \ を / に変換)
    sandboxPath = Replace(ThisWorkbook.Path & "\sandbox\", "\", "/")

    Dim r As Long: r = glb_TestLoop_STA
    ' --- テストループ ---
    Do While ws.Cells(r, 1).Value <> ""
        Dim testId As String: testId = ws.Cells(r, 1).Value

        If Left(testId, 1) = "x" Then GoTo Test_SKIP
        
        Dim url As String: url = ws.Cells(r, 2).Value
        If InStr(url, "ローカルパス") > 0 Then
             url = "file:///" & sandboxPath & Mid(url, 10, 99)
        End If
        Debug.Print "=== テスト実行: " & testId & " (" & url & ") ==="

        Dim cX As Long: cX = ws.Cells(r, 3).Value
        Dim cY As Long: cY = ws.Cells(r, 4).Value
        Dim expected As String: expected = ws.Cells(r, 5).Value
        Dim testNote As String: testNote = ws.Cells(r, 6).Value

        ' 1. 対象サイトへのナビゲーション要求
        rpaEngine.RunAction "Invoke-WebNavigation", CreateParams("Url", url)
        rpaEngine.RunAction "Wait-WebDocumentReady", CreateParams("TimeoutSec", 10)
        Sleep 3000 ' 念のための描画待機
        
        ' ==============================================================================
        ' TestId個別: 縦スクロール補正
        If InStr(testId, "Scroll") > 0 Then
            Call TestId_Type_Scroll(cY)
        End If
        ' TestId個別: iframe(固有のポップアップ消去)
        If InStr(testId, "iframe") > 0 Then
            Call TestId_Type_iframe1
        End If
        
        ' ==============================================================================
        ' 2. テスト用座標でのピッキング要求と結果の取得
        Dim responseStr As String
        responseStr = rpaEngine.RunAction("Test-WebPointReverse", CreateParams("ClientX", cX, "ClientY", cY))

         ' TestId個別: 物理フォールバック
        If InStr(testId, "物理Input") > 0 Or _
            InStr(testId, "物理Click") > 0 Then
            Call TestId_Type_PhysicalFallback(ws, r, testId, responseStr)
            GoTo Test_SKIP
        End If
        ' TestId個別: Robust DOM (SafeAction) テスト
        If InStr(testId, "RobustDOM") > 0 Then
            Call TestId_Type_RobustDOM(ws, r, testId)
            GoTo Test_SKIP
        End If
        
        ' ==============================================================================
        ' 3. 生のレスポンスをH列に記録
        ws.Cells(r, 8).Value = responseStr
        ' 4. 判定処理 (G列)
        ' JSONのエスケープ仕様に合わせて、期待値内の " を \" に置換する
        Dim expectedJson As String
        expectedJson = Replace(expected, """", "\""")

        If responseStr = "TIMEOUT" Then
            ws.Cells(r, 7).Value = "TIMEOUT"
            ws.Cells(r, 7).Font.Color = vbRed
        ElseIf InStr(1, responseStr, """selector"":""" & expectedJson & """") > 0 Then
            ws.Cells(r, 7).Value = "OK"
            ws.Cells(r, 7).Font.Color = vbBlue
        Else
            ws.Cells(r, 7).Value = "NG"
            ws.Cells(r, 7).Font.Color = vbRed
        End If
        
Test_SKIP:
        r = r + 1
    Loop
    
    MsgBox "テストが完了しました。", vbInformation
End Sub

' --- TestId: 縦スクロール補正 ---
Private Sub TestId_Type_Scroll(cY As Long)
    Dim viewHeight As Long
    viewHeight = 800    ' ブラウザ表示領域の高さ (**調整が必要**)
    
    If cY > viewHeight Then
        Dim dpr As Double: dpr = 1
        Dim dprRes As String
        dprRes = rpaEngine.RunAction("Invoke-WebScript", CreateParams("Js", "return window.devicePixelRatio || 1;"))
        dprRes = Trim(Replace(dprRes, "[RESULT]", ""))
        If IsNumeric(dprRes) Then dpr = CDbl(dprRes)

        Dim scrollAmount As Long
        scrollAmount = cY - (viewHeight / 2)
        
        Dim logicalScroll As Long
        logicalScroll = CLng(scrollAmount / dpr)
        
        ' scrollTo(絶対) ではなく scrollBy(相対) を使用する
        rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", "window.scrollBy(0, " & logicalScroll & ");")
        Sleep 800
                
        cY = cY - scrollAmount
        Debug.Print "スクロール補正完了: 移動量=" & scrollAmount & " / 新しいcY=" & cY & " / dpr=" & dpr
    End If
End Sub

' --- TestId: iframe 固有のポップアップ消去 ---
Private Sub TestId_Type_iframe1()
    ' 警告バナー（Learn more）の閉じるボタン(X)をJSでクリックして非表示にする
    Dim jsClearBanner As String
    jsClearBanner = "var btn = document.querySelector('.tox-notification__dismiss'); if(btn) btn.click();"
    rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", jsClearBanner)
    Sleep 1000  ' バナーが消えるアニメーション待ち
            
    ' 中のテキスト「Your content goes here.」も空にする
    jsClearBanner = "document.querySelector('#mce_0_ifr').contentDocument.getElementById('tinymce').innerHTML = '';"
'''    rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", jsClearBanner)
'''    rpaEngine.RunAction "Invoke-WebScript", CreateParams("Js", "document.querySelector('#mce_0_ifr').contentDocument.getElementById('tinymce').innerHTML = '';")
End Sub

' --- TestId: 物理フォールバック ---PhysicalFallback
Private Sub TestId_Type_PhysicalFallback(ws As Worksheet, r As Long, testId As String, responseStr As String)
    ' レスポンスから "boundingBox":{...} のJSON部分のみを正規表現で抽出
    Dim bboxJson As String
    Dim regEx As Object, matches As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = """boundingBox"":(\{.*?\})"
    regEx.IgnoreCase = True
    Set matches = regEx.Execute(responseStr)
    
    If matches.Count > 0 Then
        bboxJson = matches(0).subMatches(0)
        Dim fakeSelector As String
        fakeSelector = "#not_exist_" & Format(Now, "hhmmss") ' 存在しないセレクタ
        
        ' エラーを一時的に無視して物理操作の成否を捉える
        On Error Resume Next
        Dim fallbackRes As String
        
        If InStr(testId, "物理Input") > 0 Then
            fallbackRes = rpaEngine.RunAction("Set-WebTextInput", CreateParams( _
                "Selector", fakeSelector, "Value", "Fallback OK!", "BoundingBox", bboxJson))
        Else
            fallbackRes = rpaEngine.RunAction("Invoke-WebClick", CreateParams( _
                "Selector", fakeSelector, "BoundingBox", bboxJson))
        End If
        
        ' 結果をG列(7)、レスポンスをH列(8)に書き込み
        If Err.Number <> 0 Then
            ws.Cells(r, 7).Value = "NG (Fallback Error)"
            ws.Cells(r, 7).Font.Color = vbRed
            ws.Cells(r, 8).Value = Err.Description
            Err.Clear
        Else
            ws.Cells(r, 7).Value = "OK (Physical)"
            ws.Cells(r, 7).Font.Color = vbBlue
            ws.Cells(r, 8).Value = fallbackRes & " | bbox: " & bboxJson
        End If
        On Error GoTo ErrorHandler ' グローバルエラーハンドラに戻す
    Else
        ws.Cells(r, 7).Value = "NG (No BBox)"
        ws.Cells(r, 7).Font.Color = vbRed
    End If

    Exit Sub

ErrorHandler:
    MsgBox "マクロ実行エラー: " & Err.Description, vbCritical
End Sub

' --- TestId: Robust DOM (SafeAction) 検証 ---: ** サイトとの連携が必要
Private Sub TestId_Type_RobustDOM(ws As Worksheet, r As Long, testId As String)
    Dim fallbackRes As String
    Dim hintSelector As String
    Dim fuzzyXPath As String
    
    ' E列(expected)にテスト用のXPathを入れておく想定！
    fuzzyXPath = ws.Cells(r, 5).Value
    If fuzzyXPath = "" Then fuzzyXPath = "//button[contains(., '動的要素')]" ' デフォルト値
    
    On Error Resume Next
    
    ' 1. Get-WebCssSelectorHint で曖昧なXPathから最適なCSSを生成
    hintSelector = rpaEngine.RunAction("Get-WebCssSelectorHint", CreateParams("XPath", fuzzyXPath))
    hintSelector = Trim(Replace(hintSelector, "[RESULT]", ""))
    
    If hintSelector = "" Or InStr(hintSelector, "エラー") > 0 Then
        ws.Cells(r, 7).Value = "NG (Hint Error)"
        ws.Cells(r, 7).Font.Color = vbRed
        ws.Cells(r, 8).Value = Err.Description
        Err.Clear
        Exit Sub
    End If
    
    ' 2. 生成された一意のセレクタを使って、Invoke-WebSafeClick で安全にクリック
    fallbackRes = rpaEngine.RunAction("Invoke-WebSafeClick", CreateParams("Selector", hintSelector, "TimeoutSec", 10))
    
    If Err.Number <> 0 Then
        ws.Cells(r, 7).Value = "NG (SafeClick Error)"
        ws.Cells(r, 7).Font.Color = vbRed
        ws.Cells(r, 8).Value = Err.Description
        Err.Clear
    Else
        ws.Cells(r, 7).Value = "OK (Robust)"
        ws.Cells(r, 7).Font.Color = vbBlue
        ws.Cells(r, 8).Value = fallbackRes & " | Generated: " & hintSelector
    End If
End Sub

' ==============================================================================
'   RPA-Engine／テスト (88)
' ==============================================================================
Private Sub Test_DevTools_88(targetUrl As String)
    UserForm_Picker.TextUrl.Text = ""
    UserForm_Picker.lblStatus.Caption = "任意のURLを選択（入力）して、[移動]ボタンを押してください。"
End Sub
' ==============================================================================
'   RPA-Engine／テスト (99)
' ==============================================================================
Private Sub Test_DevTools_99()
    'xx
End Sub
