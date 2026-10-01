# ------------------------------------------------------------------------------#●#✘ 
<#
  [デスクトップ操作モジュール]： Ps_Engine_Core を利用したOSネイティブ操作(.NET UIAutomationとWin32 API)を行い、
  ダイアログ操作や物理フォールバックとして機能する。
#>
# ------------------------------------------------------------------------------

# --- 外部アプリケーションウィンドウのアクティブ化 ---
# Win32 APIを用いて指定した外部ウィンドウを最前面へ引き上げる。
function Switch-AppWindow {
    param ([string]$Name)

    $func = $MyInvocation.MyCommand.Name

    # ウィンドウの検索： 完全一致および部分一致によるウィンドウ検索を行う。
    $hWnd = [Win32Api.Win32Utils]::FindWindow($null, $Name)
    if ($hWnd -eq [IntPtr]::Zero -or $hWnd -eq $null) {
        # FindWindowで完全一致取得できなかった場合、プロセス一覧からタイトルを部分一致検索してハンドルを取得する。
        $proc = Get-Process | Where-Object { $_.MainWindowTitle -like "*$Name*" } | Select-Object -First 1
        if ($proc) { $hWnd = $proc.MainWindowHandle }
    }

    # ウィンドウの前面化： 取得したハンドルを用いてウィンドウを元のサイズで復元し最前面化する。
    if ($hWnd -and $hWnd -ne [IntPtr]::Zero) {
        # ウィンドウを元のサイズで復元・アクティブ化(9 = SW_RESTORE)し、Zオーダーの最前面へ引き上げる。
        [Win32Api.Win32Utils]::ShowWindow($hWnd, 9) | Out-Null
        [Win32Api.Win32Utils]::SetForegroundWindow($hWnd) | Out-Null
        return "Focused: $Name"
    }
    throw (New-EngineException -Func $func -Type "未発見" -Message "指定されたウィンドウが見つかりません" -Details $Name)
}

# --- クリック処理専用ヘルパー ---
# 要素のバックグラウンドクリック、または物理クリックを実行する（内部関数）。
function Invoke-UiaClickInternal {
    param ($Element, $Mode, $Func)
    
    if ($Mode -eq "Safe") {
        # 安全モード（物理クリック）： 要素をフォーカスし、Spaceキーを物理送信してクリックを代替する。
        try { $Element.SetFocus() } catch {}
        Start-Sleep -Milliseconds 100
        $wshell = New-Object -ComObject WScript.Shell
        $wshell.SendKeys(" ")
    } else {
        # パターンモード（バックグラウンドクリック）： InvokePatternが利用可能な場合はバックグラウンドでクリックを発火する。
        $invokePattern = $null
        
        # InvokePattern(ボタン押下等のイベント発火API)をサポートしているか確認し、可能ならバックグラウンドで発火させる。
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$invokePattern)) {
            $invokePattern.Invoke()
        } else {
            Write-DebugLog -Message "[$Func] 警告: InvokePattern非対応のため、Safeモード(物理クリック)へ自動フォールバックします。" -Level Warning

            try { $Element.SetFocus() } catch {}
            Start-Sleep -Milliseconds 100
            $wshell = New-Object -ComObject WScript.Shell
            $wshell.SendKeys(" ")
        }
    }
}

# --- 入力処理専用ヘルパー ---
# 要素への値の直接書き込み、または物理キー送信を実行する（内部関数）。
function Invoke-UiaInputInternal {
    param ($Element, $Value, $Mode, $Func)
    
    if ($Mode -eq "Safe") {
        # 安全モード（物理入力）： 要素にフォーカスを当て、キーボードストロークとして直接文字列を送信する。
        try { $Element.SetFocus() } catch {}
        Start-Sleep -Milliseconds 100
        $wshell = New-Object -ComObject WScript.Shell
        $wshell.SendKeys($Value)
    } else {
        # パターンモード（バックグラウンド入力）：
        $valuePattern = $null

        # ValuePattern(値の直接書き込みAPI)をサポートしているか確認し、可能なら直接値をセットする。
        if ($Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
            $valuePattern.SetValue($Value)
        } else {
            # ComboBox等の場合、内部にある子要素(Edit)がValuePatternを持っているか確認する。
            $childCondition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
            $editChild = $Element.FindFirst([System.Windows.Automation.TreeScope]::Children, $childCondition)
            
            if ($editChild -and $editChild.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
                $valuePattern.SetValue($Value)
            } else {
                # それでもダメなら、エラーで落とさず自動でSafeモードへ移行する。
                Write-DebugLog -Message "[$Func] 警告: ValuePattern非対応のため、Safeモード(物理入力)へ自動フォールバックします。" -Level Warning
                try { 
                    if ($editChild) { $editChild.SetFocus() } else { $Element.SetFocus() } 
                } catch {} 
                
                Start-Sleep -Milliseconds 100
                $wshell = New-Object -ComObject WScript.Shell
                # 既存の入力値を確実に取り除くため、全選択して削除してから送信する。
                $wshell.SendKeys("^a") # 全選択
                Start-Sleep -Milliseconds 50
                $wshell.SendKeys("{DELETE}") # 削除
                Start-Sleep -Milliseconds 50
                $wshell.SendKeys($Value) # 入力
            }
        }
    }
}

# --- 値取得専用ヘルパー ---
# 要素から ValuePattern または Name を用いて値を取得する（内部関数）。
function Get-UiaValueInternal {
    param ($Element)
    
    $valuePattern = $null
    # ValuePatternが存在する場合はそのCurrent.Valueを返す。
    if ($Element.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
        return $valuePattern.Current.Value
    }
    
    # GetValueもComboBoxの子要素(Edit)まで探しに行く。
    $childCondition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ControlTypeProperty, [System.Windows.Automation.ControlType]::Edit)
    $editChild = $Element.FindFirst([System.Windows.Automation.TreeScope]::Children, $childCondition)
    
    if ($editChild -and $editChild.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
        return $valuePattern.Current.Value
    }
    return $Element.Current.Name
}

# --- UIAを利用したデスクトップアプリの要素操作 ---
# UIAutomationを用い、バックグラウンドパターンまたは物理キー送信でOS要素を操作する。
function Invoke-UiaAction {
    param (
        [Parameter(Mandatory=$true)][string]$Action,
        [Parameter(Mandatory=$true)][string]$Name = "",
        [string]$AutomationId = "",
        [string]$Value = "",
        [string]$Mode = "Pattern",
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec,
        [int]$PollIntervalMs = $global:CONFIG.PollIntervalMs
    )

    $func = $MyInvocation.MyCommand.Name

    try {
        # 検索条件の構築： AutomationIdを優先して検索条件を構築する。
        $condition = if (![string]::IsNullOrEmpty($AutomationId)) {
            New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::AutomationIdProperty, $AutomationId
            )
        } else {
            New-Object System.Windows.Automation.PropertyCondition(
                [System.Windows.Automation.AutomationElement]::NameProperty, $Name
            )
        }

        # UIA要素の探索： タイムアウト付きで最前面ウィンドウへ動的に追従し、UIA要素を探索する。
        $element = $null
        $sw = [System.Diagnostics.Stopwatch]::StartNew()

        while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
            # 現在最前面にあるウィンドウのハンドルを取得する。
            $hwnd = [Win32Api.Win32Utils]::GetForegroundWindow()
            if ($hwnd -ne [IntPtr]::Zero) {
                try {
                    $root = [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
                    $element = $root.FindFirst([System.Windows.Automation.TreeScope]::Subtree, $condition)
                    if ($element) { break }
                } catch {}
            }
            Start-Sleep -Milliseconds $PollIntervalMs
        }

        if (-not $element) { throw (New-EngineException -Func $func -Type "未発見" -Message "指定されたUIA要素が見つかりません" -Details "Timeout: ${TimeoutSec}s") }

        # アクションの実行： アクション種別に応じてヘルパー関数へルーティングする。
        switch ($Action) {
            "Click"    { Invoke-UiaClickInternal -Element $element -Mode $Mode -Func $func }
            "Input"    { Invoke-UiaInputInternal -Element $element -Value $Value -Mode $Mode -Func $func }
            "GetText"  { return $element.Current.Name }
            "GetValue" { return Get-UiaValueInternal -Element $element }
            default    { throw (New-EngineException -Func $func -Type "引数エラー" -Message "未対応のUIAアクションが指定されました" -Details $Action) }
        }
        return "Action Completed: $Action"

    } catch {
        # 既に New-EngineException で生成されたエラーならそのまま throw する。
        if ($_.Exception.Message -match "^\[.*?\] \[.*?\]:") { throw $_ }
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "UIA操作中に予期せぬエラーが発生しました" -Details $_.Exception.Message)
    }
}

# --- 「名前を付けて保存」ダイアログの捕捉および保存の実行 ---
# 「名前を付けて保存」ダイアログを捕捉し、クリップボード経由でパスを入力・保存する。
function Invoke-UiaSafeSaveAs {
    param(
        [Parameter(Mandatory=$true)][string]$FilePath,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    $dialogNames = @("名前を付けて保存", "Save As", "保存", "Save")
    $hwnd = [IntPtr]::Zero
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    # Win32APIによる安全なハンドル取得： クラス名等を駆使し、OS標準のダイアログを探索する。
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        foreach ($name in $dialogNames) {
            <#
              PowerShellの $null はWin32API(C#)に渡る際 "" (空文字) に変換されてしまうため、
              [NullString]::Value を使用するか、標準ダイアログクラス "#32770" を明示指定する。
            #>
            
            # パターンA： 標準のダイアログクラス名(#32770)で検索する。
            $hwnd = [Win32Api.Win32Utils]::FindWindow("#32770", $name)
            
            if ($hwnd -eq [IntPtr]::Zero -or $hwnd -eq $null) {
                # パターンB： クラス名問わず、完全な null ポインタとして検索する。
                $hwnd = [Win32Api.Win32Utils]::FindWindow([NullString]::Value, $name)
            }

            if ($hwnd -ne [IntPtr]::Zero -and $hwnd -ne $null) { break }
        }
        if ($hwnd -ne [IntPtr]::Zero -and $hwnd -ne $null) { break }
        Start-Sleep -Milliseconds 200
    }

    if ($hwnd -eq [IntPtr]::Zero -or $hwnd -eq $null) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "保存ダイアログのウィンドウが見つかりませんでした")
    }

    # ウィンドウの最前面への強制引き上げ：
    try {
        [Win32Api.Win32Utils]::ShowWindow($hwnd, 9) | Out-Null
        [Win32Api.Win32Utils]::SetForegroundWindow($hwnd) | Out-Null
        Start-Sleep -Milliseconds 300
    } catch {}

    # UIA要素の生成： 捕捉した安全なダイアログのハンドルを起点とする。
    $dialog = $null
    try {
        $dialog = [System.Windows.Automation.AutomationElement]::FromHandle($hwnd)
    } catch {
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "ハンドルからUIA要素への変換に失敗しました")
    }

    # ファイル名入力欄の取得： AutomationId=1001を指定して入力欄を取得する。
    $cndEdit = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::AutomationIdProperty, "1001"
    )
    $editBox = $dialog.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $cndEdit)

    if ($null -eq $editBox) {
        Write-DebugLog -Message "[$func] 警告: 入力欄(1001)が未発見。デフォルトフォーカスを利用します。" -Level Warning
    } else {
        try {
            # 確実に入力を行うため、対象のテキストエリアにシステムフォーカスを移動する。
            $editBox.SetFocus()
            Start-Sleep -Milliseconds 200
        } catch {}
    }

    # クリップボードと物理キーによる安全な操作： SendKeysによる文字化けを防ぐためクリップボード経由でパスをペーストする。
    try {
        Write-DebugLog -Message "[$func] 情報: 物理キー(SendKeys)で保存を実行します" -Level Info

        $wshell = New-Object -ComObject WScript.Shell

        # 既存のテキストを全選択して削除 (Ctrl+A -> Delete)を実行する。
        $wshell.SendKeys("^a")
        Start-Sleep -Milliseconds 100
        $wshell.SendKeys("{DELETE}")
        Start-Sleep -Milliseconds 100

        <# 
          クリップボード経由でパスを貼り付ける。
          SendKeysでの直接タイピングは、日本語IMEの干渉やPC負荷による文字抜け(パス欠損)のリスクが高いため、クリップボード経由での確実なペースト(Ctrl+V)を採用する。
        #>
        Set-Clipboard -Value $FilePath
        Start-Sleep -Milliseconds 100
        # 貼り付けコマンド(Ctrl+V)を送信する。
        $wshell.SendKeys("^v")
        Start-Sleep -Milliseconds 300

        # Enterキーで保存を実行する。
        $wshell.SendKeys("{ENTER}")

    } catch {
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "パスの入力または保存の実行に失敗しました" -Details $_.Exception.Message)
    }

    # 保存完了待機： ファイルロックが解除されるまで待機する。
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        if (Test-Path $FilePath) {
            try {
                $stream = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'None')
                $stream.Close()

                return $FilePath
            } catch {}
        }
        Start-Sleep -Milliseconds 300
    }
    throw (New-EngineException -Func $func -Type "Timeout" -Message "ファイル保存の完了確認ができませんでした" -Details $FilePath)
}

# --- アクティブウィンドウへの物理キー送信 ---
function Invoke-DesktopSendKeys {
    param (
        [Parameter(Mandatory=$true)][string]$Keys,
        [int]$WaitMs = 300
    )

    $func = $MyInvocation.MyCommand.Name

    try {
        # 画面のフォーカスが安定するまで少し待機する。
        Start-Sleep -Milliseconds $WaitMs
        
        # WScript.Shellオブジェクトを生成し、引数で渡されたキー操作文字列をOSへ送信する。
        $wshell = New-Object -ComObject WScript.Shell
        $wshell.SendKeys($Keys)
        
        Start-Sleep -Milliseconds 300
        return "SendKeys Completed: $Keys"

    } catch {
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "キーボード送信に失敗しました" -Details $_.Exception.Message)
    }
}

# --- アクティブウィンドウの中央への物理クリック実行 ---
function Invoke-DesktopCenterClick {
    param ([int]$WaitMs = 500)

    $func = $MyInvocation.MyCommand.Name

    try {
        # 画面のフォーカスが安定するまで少し待機する。
        Start-Sleep -Milliseconds $WaitMs
        
        # 最前面のウィンドウハンドルの取得：
        $hwnd = [Win32Api.Win32Utils]::GetForegroundWindow()
        if ($hwnd -eq [IntPtr]::Zero) {
            throw "最前面のウィンドウが取得できません"
        }
        
        # ウィンドウの矩形領域の取得： 取得したハンドルから、ウィンドウの物理的な画面上の位置とサイズ(RECT)を取得する。
        $rect = New-Object Win32Api.RECT
        $res = [Win32Api.Win32Utils]::GetWindowRect($hwnd, [ref]$rect)
        if (-not $res) {
            throw "ウィンドウ領域の取得に失敗しました"
        }
        
        # 画面中央の座標の計算： ウィンドウの左上座標に、幅と高さの半分を加算して中心座標を算出する。
        $centerX = [int]($rect.Left + (($rect.Right - $rect.Left) / 2))
        $centerY = [int]($rect.Top + (($rect.Bottom - $rect.Top) / 2))
        
        # マウスカーソルの中央への移動： 算出した中心座標へマウスポインタを瞬間移動させる。
        [Win32Api.Win32Utils]::SetCursorPos($centerX, $centerY) | Out-Null
        
        # 物理クリックの発火： LeftDown(0x0002)およびLeftUp(0x0004)を送信する。
        [Win32Api.Win32Utils]::mouse_event(0x0002, 0, 0, 0, 0)
        [Win32Api.Win32Utils]::mouse_event(0x0004, 0, 0, 0, 0)
        
        return "Clicked Center: X=$centerX, Y=$centerY"

    } catch {
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "画面中央の物理クリックに失敗しました" -Details $_.Exception.Message)
    }
}

# --- 指定座標への物理マウスクリック実行 ---
# 指定されたX/Y座標（OS絶対スクリーン座標）に対して物理マウスクリックを実行する。
<# 
  注意: ここで指定する $X, $Y は、ブラウザの論理座標ではなく、DPR(画面拡大率)を加味した「OSの絶対スクリーン座標(物理ピクセル)」である必要がある。
#>
function Invoke-DesktopCoordinateClick {
    param (
        [Parameter(Mandatory=$true)][int]$X,
        [Parameter(Mandatory=$true)][int]$Y,
        [int]$WaitMs = 100
    )

    $func = $MyInvocation.MyCommand.Name

    try {
        # マウスカーソルの移動： 物理ピクセル座標へマウスカーソルを瞬間移動させる。
        [Win32Api.Win32Utils]::SetCursorPos($X, $Y) | Out-Null
        Start-Sleep -Milliseconds $WaitMs
        
        # 物理クリックの発火： LeftDown(0x0002)およびLeftUp(0x0004)を送信する。
        [Win32Api.Win32Utils]::mouse_event(0x0002, 0, 0, 0, 0)
        [Win32Api.Win32Utils]::mouse_event(0x0004, 0, 0, 0, 0)
        
        Write-DebugLog -Message "[$func] 物理クリック実行: X=$X, Y=$Y" -Level Info
        return $true

    } catch {
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "物理クリックに失敗しました" -Details $_.Exception.Message)
    }
}
