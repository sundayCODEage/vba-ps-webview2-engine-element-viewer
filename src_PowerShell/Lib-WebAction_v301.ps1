# ------------------------------------------------------------------------------#●#✘ 
<#
  [Web標準操作モジュール]： DOMベースの要素探索、待機、クリック、入力などの共通アクションを定義する。
#>
# ------------------------------------------------------------------------------

# --- BoundingBox(絶対論理座標)をOSの絶対スクリーン座標へ変換 ---
# ブラウザ内の論理座標を、Win32API等で操作可能なOS絶対スクリーン座標(物理ピクセル)に変換する。
function Convert-BoundingBoxToScreenPoint {
    param ([Parameter(Mandatory=$true)][object]$BoundingBox)

    # 1. スクロール量の取得： ウィンドウの現在のX/Yスクロール量をJSONで取得する。
    $scrollX = 0; $scrollY = 0
    try {
        # windowのスクロール量をX/YのJSONオブジェクトとして取得する。
        $scrJs = "return JSON.stringify({x: window.scrollX || 0, y: window.scrollY || 0});"
        $scrRes = Invoke-WebScript -Js $scrJs -Retries 1
        $scrObj = $scrRes | ConvertFrom-Json
        if ($scrObj) { $scrollX = $scrObj.x; $scrollY = $scrObj.y }
    } catch {}

    # 2. クライアント論理座標の算出： 絶対座標からスクロール量を引き、要素の中心を算出する。
    $clientLogicalX = ($BoundingBox.x - $scrollX) + ($BoundingBox.width / 2)
    $clientLogicalY = ($BoundingBox.y - $scrollY) + ($BoundingBox.height / 2)

    # 3. クライアント物理座標への変換： DPRを掛けて物理ピクセル座標へ変換する。
    $dprStr = Invoke-WebScript -Js "return window.devicePixelRatio || 1;"
    $dpr = if ([double]::TryParse($dprStr, [ref]$null)) { [double]$dprStr } else { 1.0 }
    $clientX = [int][Math]::Round($clientLogicalX * $dpr)
    $clientY = [int][Math]::Round($clientLogicalY * $dpr)

    # 4. OSの絶対スクリーン座標への変換： PointToScreenを用いてブラウザ相対座標をOS絶対座標へ変換する。
    if (-not ('System.Drawing.Point' -as [type])) { Add-Type -AssemblyName System.Drawing }
    $clientPt = New-Object System.Drawing.Point($clientX, $clientY)
    return $global:WebViewCtrl.PointToScreen($clientPt)
}

# --- 指定URLへのナビゲーション実行 ---
# アクティブなWebView2インスタンスに対して、指定URLへのページ遷移を実行する。
function Invoke-WebNavigation {
    param ([Parameter(Mandatory=$true)][string]$Url)

    $func = $MyInvocation.MyCommand.Name

    try {
        # アクティブWebViewに対してNavigateを実行する。
        $activeWv = Get-ActiveWebView
        $activeWv.CoreWebView2.Navigate($Url)
        return "Navigating to $Url"
    } catch {
        throw (New-EngineException -Func $func -Type "ネイティブエラー" -Message "指定されたURLへのナビゲーションに失敗しました" -Details $_.Exception.Message)
    }
}

# --- ページ読み込み完了（iframe含む完全ロード）の待機 ---
# DOMの readyState=complete を全iframe含めて再帰的に待機する。
function Wait-WebPageLoad {
    param ([int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec)

    $func = $MyInvocation.MyCommand.Name

    # readyState=completeおよび全iframeのDOM完成の再帰チェック：
    $condition = {
        try {
            $js = @"
                function isLoaded(win) {
                    try {
                        // ドキュメントの状態が完了(complete)でなければfalseを返す。
                        if (win.document.readyState !== 'complete') return false;

                        // iframeのDOMがまだ構築中のケースに対応する。
                        let frames = win.frames;
                        for (let i = 0; i < frames.length; i++) {
                            try {
                                // 子フレームに対しても再帰的に状態をチェックする。
                                if (!isLoaded(frames[i])) return false;
                            } catch(e) {
                                // アクセス不可(CORS)は無視する。
                            }
                        }
                        return true;
                    } catch(e) {
                        return false;
                    }
                }
                return isLoaded(window);
"@

            # CDPモード時におけるネイティブ実行を強制する。
            $res = Invoke-WebView2NativeScript -Js $js

            if ($res -eq $true) {
                # DOM安定化のため追加待機を実行する。
                Start-Sleep -Milliseconds 300
                return $true
            }
        } catch {}
        return $false
    }
    $errMsg = "[$func] タイムアウト: ページまたは iframe のロード未完了"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null

    return $true
}

# --- 画面全体の読み込みステータス（complete）待機 ---
# 画面全体の読み込みステータス完了を高速に待機する。
function Wait-WebDocumentReady {
    param ([int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec)

    $func = $MyInvocation.MyCommand.Name

    $js = @"
        const res = utilFindInFrames(window, function(win) {
            try {
                return (win.document.readyState === 'complete');
            } catch(e) {
                return null;
            }
        });
        return res === true;
"@

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $isReady = $false

    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        # 修正後のラッパー構造を介して安全にBoolean（true/false）を受け取る。
        $state = Invoke-WebScript -Js $js -Retries 1
        
        if ($state -eq $true -or $state -eq "true") {
            $isReady = $true
            break
        }
        # UIをフリーズさせないためにイベントを処理する。
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 300
    }

    if (-not $isReady) {
        throw (New-EngineException -Func $func -Type "Timeout" -Message "Document ReadyState が制限時間内に complete になりませんでした" -Details "Timeout: ${TimeoutSec}s")
    }
}

# --- URLの部分一致待機 ---
# URLに指定文字列が含まれるまで待機する。
function Wait-WebUrlContains {
    param (
        [Parameter(Mandatory=$true)][string]$Substring,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    # URLへの指定文字列包含待機： 現在のURLに指定文字列が含まれるか部分一致検索する。
    $condition = {
        $webview = Get-ActiveWebView
        if ($webview -and $webview.Source) {
            # 現在のURL文字列を取得し、ワイルドカード検索で部分一致するか判定する。
            $url = $webview.Source.ToString()
            if ($url -like "*$Substring*") { return $true }
        }
        return $false
    }
    $errMsg = "[$func] タイムアウト: URL不一致 ($Substring)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -PollIntervalMs 300 -TimeoutMessage $errMsg | Out-Null

    return "URL matched: $(Get-WebUrl)"
}

# --- ページタイトルの部分一致待機 ---
# タイトルに指定文字列が含まれるまで待機する。
function Wait-WebTitleContains {
    param (
        [Parameter(Mandatory=$true)][string]$Substring,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    # タイトルへの指定文字列包含待機：
    $condition = {
        $webview = Get-ActiveWebView
        if ($webview) {
            # WebView2コントロールのプロパティから直接タイトルを取得し、大文字小文字を区別せずに一致判定する。
            $title = $webview.CoreWebView2.DocumentTitle
            if ($title -and $title.Trim().IndexOf($Substring, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                return $true
            }
        }
        return $false
    }
    $errMsg = "[$func] タイムアウト: タイトル不一致 ($Substring)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -PollIntervalMs 300 -TimeoutMessage $errMsg | Out-Null

    return "Title matched: $(Get-WebTitle)"
}

# --- 指定要素の出現および可視化待機 ---
# 指定要素がDOM上に出現し、かつ画面上に可視化されるまで待機する。
function Wait-WebElement {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    # Base64エンコードを利用してサニタイズする。
    $selB64 = ConvertTo-JsSafeBase64 -Text $Selector

    # DOM出現および可視判定： display/visibility/opacity/rectを利用して要素の状態を判定する。
    $condition = {
        try {
            $js = @"
                const selector = decodeURIComponent(escape(atob('$selB64')));
                const found = utilFindInFrames(window, function(win) {
                    try {
                        // 共通関数を利用し、Shadow DOM内を含めて要素を探索する。
                        const el = deepQuerySelector(selector, win.document);
                        if (el) {
                            // 要素が存在する場合、スタイルとサイズを取得して可視状態を判定する。
                            const style = win.getComputedStyle(el);
                            const rect = el.getBoundingClientRect();
                            // 要素が存在し、かつ可視状態なら true を返す。
                            return (style.display !== 'none' 
                                    && style.visibility !== 'hidden'
                                    && style.opacity !== '0'
                                    && rect.width > 0 && rect.height > 0);
                        }
                    } catch(e) {}
                    return null;
                });
                return found === true;
"@

            $res = Invoke-WebScript -Js $js
            if ($res) { return $true }
            
        } catch {}
        return $false
    }
    $errMsg = "[$func] タイムアウト: 要素の未出現または非表示 ($Selector)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null

    return $true
}

# --- iframe内要素の出現および可視化待機 ---
# 指定したiframe内の要素が出現・可視化されるまで待機する。
function Wait-WebElementInFrame {
    param (
        [Parameter(Mandatory=$true)][string]$FrameSelector,
        [Parameter(Mandatory=$true)][string]$ElementSelector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    # Base64エンコードを利用してサニタイズする。
    $frameB64 = ConvertTo-JsSafeBase64 -Text $FrameSelector
    $elementB64 = ConvertTo-JsSafeBase64 -Text $ElementSelector

    # iframe.contentDocumentを用いた直接探索： iframe内部へ直接アクセスし要素の可視性を判定する。
    $condition = {
        try {
            $js = @"
                const frmSel = decodeURIComponent(escape(atob('$frameB64')));
                const elSel = decodeURIComponent(escape(atob('$elementB64')));

                // 親ドキュメントからiframe要素を特定する。
                const frm = document.querySelector(frmSel);
                if (!frm || !frm.contentDocument) return 'frame_not_found';
                // iframe内部のドキュメントから目的の要素を特定する。
                const el = frm.contentDocument.querySelector(elSel);
                if (!el) return 'not_found';
                // iframe内部のwindowコンテキストでスタイルを取得し、可視状態を判定する。
                const style = frm.contentWindow.getComputedStyle(el);
                const rect = el.getBoundingClientRect();
                const isVisible = (style.display !== 'none'
                                 && style.visibility !== 'hidden'
                                 && style.opacity !== '0'
                                 && rect.width > 0 && rect.height > 0);
                return isVisible ? 'visible' : 'hidden';
"@

            $res = Invoke-WebScript -Js $js
            if ($res -eq "visible") { return $true }
        } catch {}
        return $false
    }
    $errMsg = "[$func] タイムアウト: iframe内要素の未出現 ($ElementSelector)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null

    return $true
}

# --- iframe内要素のクリック実行 ---
# 指定したiframe内の要素をスクロールしてクリックする。
function Invoke-WebClickInFrame {
    param (
        [Parameter(Mandatory=$true)][string]$FrameSelector,
        [Parameter(Mandatory=$true)][string]$ElementSelector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    # Base64エンコードを利用してサニタイズする。
    $frameB64 = ConvertTo-JsSafeBase64 -Text $FrameSelector
    $elementB64 = ConvertTo-JsSafeBase64 -Text $ElementSelector

    # 対象要素がiframe内に出現し可視状態になるまで待機する。
    Wait-WebElementInFrame -FrameSelector $FrameSelector -ElementSelector $ElementSelector -TimeoutSec $TimeoutSec | Out-Null

    $js = @"
        const frmSel = decodeURIComponent(escape(atob('$frameB64')));
        const elSel = decodeURIComponent(escape(atob('$elementB64')));

        const frm = document.querySelector(frmSel);
        if (frm && frm.contentDocument) {
            const el = frm.contentDocument.querySelector(elSel);
            if (el) {
                // iframe内で要素を中央へスクロールし、クリックする。
                el.scrollIntoView({block: 'center', inline: 'center'});
                el.click();
                return true;
            }
        }
        return false;
"@

    $res = Invoke-WebScript -Js $js
    if (-not $res) { throw (New-EngineException -Func $func -Type "未発見" -Message "iframe内でのクリック実行に失敗しました" -Details $ElementSelector) }
}

# --- 指定要素の非表示またはDOM削除待機 ---
# 指定要素が非表示になる、またはDOMから消滅するまで待機する。
function Wait-WebElementInvisible {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    # Base64エンコードを利用してサニタイズする。
    $selB64 = ConvertTo-JsSafeBase64 -Text $Selector

    # 多段iframe対応および可視判定： display/visibility/opacity/rectを利用して要素の状態を判定する。
    $condition = {
        $js = @"
            const selector = decodeURIComponent(escape(atob('$selB64')));
            const found = utilFindInFrames(window, function(win) {
                try {
                    const el = deepQuerySelector(selector, win.document);
                    if (el) {
                        const style = win.getComputedStyle(el);
                        const rect = el.getBoundingClientRect();
                        const isVisible = (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && rect.width > 0 && rect.height > 0);
                        if (isVisible) return true; // まだ表示されている
                    }
                } catch(e) {}
                return null;
            });
            // どこにも無いか、あっても非表示なら 'hidden' を返す。
            return found === true ? 'visible' : 'hidden';
"@

        $res = Invoke-WebScript -Js $js
        if ($res -eq "hidden") { return $true }
        return $false
    }
    $errMsg = "[$func] タイムアウト: 要素が非表示になりません ($Selector)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null

    return $true
}

# --- 画面のローディングマスク解除待機（汎用） ---
# 業務システム特有のローディングマスク（透過レイヤー）の解除を待機する。
function Wait-WebScreenUnlock {
    param ([int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec)

    $func = $MyInvocation.MyCommand.Name

    <#
      ※ 汎用マスク判定ロジック
      z-indexが100以上で、画面の80%以上を覆っており、透明でないdiv要素をローディングマスクとみなす。
      また、iframe内のマスクも再帰的に検知する。
    #>
    $js = @"
        function checkMask(doc, win) {
            // 画面上のすべてのdiv要素を取得する。
            const divs = doc.querySelectorAll('div');
            for (let d of divs) {
                const s = win.getComputedStyle(d);
                // s.zIndexが "auto" や空文字だった場合に "0" としてパースさせる。
                const z = parseInt(s.zIndex || "0", 10);
                // 汎用判定：一般的なローディングスピナーやオーバーレイのCSS特性を満たすか判定する。
                if (!isNaN(z) && z >= 100 && (s.position === 'fixed' || s.position === 'absolute') &&
                    d.offsetWidth >= win.innerWidth * 0.8 && d.offsetHeight >= win.innerHeight * 0.8 &&
                    s.display !== 'none' && s.visibility !== 'hidden' && s.opacity !== '0') {
                    return true;
                }
            }
            
            // iframe / frame の中も再帰的に探査する。
            const frames = doc.querySelectorAll('iframe, frame');
            for (let f of frames) {
                try {
                    if (f.contentDocument && f.contentWindow) {
                        if (checkMask(f.contentDocument, f.contentWindow)) return true;
                    }
                } catch(e) { 
                    // クロスドメイン(CORS)のセキュリティエラーは安全に無視する。
                }
            }
            return false;
        }
        return checkMask(document, window);
"@

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $wasBlocked = $false
    $initialGraceMs = 500 # ボタン押下後のマスク描画猶予時間(ms)

    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        try {
            $rawResult = Invoke-WebScript -Js $js

            # 配列返却時は末尾要素を抽出する。
            if ($rawResult -is [array] -and $rawResult.Count -gt 0) {
                $isBlocked = $rawResult[-1]
            } else {
                $isBlocked = $rawResult
            }

            # 文字列・Bool値の判定ロジック：
            $blockedBool = ($isBlocked -eq $true -or $isBlocked -eq "true")

            if ($blockedBool) {
                # マスク出現を検知する。
                $wasBlocked = $true
            }
            else {
                # マスク非表示時の完了判定： 一度でもマスクを検知した、または描画猶予時間を過ぎていれば完了とする。
                if ($wasBlocked -or $sw.ElapsedMilliseconds -gt $initialGraceMs) {
                    if ($wasBlocked) {
                        Write-DebugLog -Message "[$func] 情報: マスク解除を確認しました" -Level Info
                    }
                    return "Screen Unlocked"
                }
            }
        } catch {
            # DOMアクセスエラー等は無視してリトライする。
        }

        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 200
    }
    throw (New-EngineException -Func $func -Type "Timeout" -Message "画面のマスク解除がタイムアウトしました。誤検知またはシステム遅延の可能性があります。" -Details "Timeout: ${TimeoutSec}s")
}

# --- 指定要素のクリック実行（物理クリック・フォールバック対応版） ---
# 多段iframeを透過的に探索し、対象要素をクリックする。失敗時は物理クリックへフォールバックする。
function Invoke-WebClick {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec,
        [string]$OuterHtml = "",
        [string]$BoundingBox = ""  # ピッカーからの座標データを受け取る。
    )

    $func = $MyInvocation.MyCommand.Name
    # Base64エンコードを利用してサニタイズする。
    $selB64 = ConvertTo-JsSafeBase64 -Text $Selector

    # 要素出現を待機する。（エラーを握り潰してフォールバックへ流すため Try-Catch を適用する）
    try { Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null } catch {}

    # 多段iframeの再帰探索およびJSクリック：
    $js = @"
        const selector = decodeURIComponent(escape(atob('$selB64')));
        const found = utilFindInFrames(window, function(win) {
            try {
                const el = deepQuerySelector(selector, win.document);
                if (el) {
                    // 要素を画面の中央へスクロールし、クリックを実行する。
                    el.scrollIntoView({block: 'center', inline: 'center'});
                    el.click();
                    return true;
                }
            } catch(e) {}
            return null;
        });
        return found === true;
"@

    $res = Invoke-WebScript -Js $js

    # JSクリック失敗時の物理クリックフォールバック：
    if (-not $res -and -not [string]::IsNullOrWhiteSpace($BoundingBox)) {
        try {
            $bbox = $BoundingBox | ConvertFrom-Json
            if ($bbox.width -gt 0 -and $bbox.height -gt 0) {
                # 論理座標をOSの絶対スクリーン座標へ変換する。
                $screenPt = Convert-BoundingBoxToScreenPoint -BoundingBox $bbox
                $targetX = $screenPt.X
                $targetY = $screenPt.Y

                Write-DebugLog -Message "[$func] 情報: JSクリック失敗。BoundingBox座標(X:$targetX, Y:$targetY)へ物理クリックを試行します。" -Level Info
                
                # Lib-DesktopUIA の物理クリック関数を呼び出す。
                if (Get-Command -Name "Invoke-DesktopCoordinateClick" -ErrorAction SilentlyContinue) {
                    Invoke-DesktopCoordinateClick -X $targetX -Y $targetY
                    return "[FALLBACK_PHYSICAL] Click Completed: $Selector"
                }
            }
        } catch {
            Write-DebugLog -Message "[$func] エラー: 物理クリックの実行中に例外が発生しました。" -Level Error
        }
    }

    if (-not $res) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "クリック実行に失敗しました" -Details $Selector)
    }
    return "JS Click Completed: $Selector"
}

# --- テキストボックスへの値入力（物理入力・フォールバック対応版） ---
# テキストボックスに値を入力し、input / change イベントを発火する。失敗時は物理キー送信へフォールバックする。
function Set-WebTextInput {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [Parameter(Mandatory=$true)][string]$Value,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec,
        [string]$OuterHtml = "",
        [string]$BoundingBox = ""  # ピッカーからの座標データを受け取る。
    )

    $func = $MyInvocation.MyCommand.Name
    # Base64エンコードを利用してサニタイズする。
    $selB64 = ConvertTo-JsSafeBase64 -Text $Selector
    $valB64 = ConvertTo-JsSafeBase64 -Text $Value

    try { Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null } catch {}

    $js = @"
        const selector = decodeURIComponent(escape(atob('$selB64')));
        const val = decodeURIComponent(escape(atob('$valB64')));
        const found = utilFindInFrames(window, function(win) {
            try {
                const el = deepQuerySelector(selector, win.document);
                if (el) {
                    // 対象要素にフォーカスを当てる。
                    el.focus();
                    // DOMのプロパティとして値を直接書き込む。
                    el.value = val;
                    // ReactやVueなどのモダンフレームワークに変更を認識させるため、イベントを強制発火させる。
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    return true;
                }
            } catch(e) {}
            return null;
        });
        return found === true;
"@

    $res = Invoke-WebScript -Js $js

    # JS入力失敗時の物理フォーカス＆キーボードフォールバック：
    if (-not $res -and -not [string]::IsNullOrWhiteSpace($BoundingBox)) {
        try {
            $bbox = $BoundingBox | ConvertFrom-Json
            if ($bbox.width -gt 0 -and $bbox.height -gt 0) {
                # 論理座標をOSの絶対スクリーン座標へ変換する。
                $screenPt = Convert-BoundingBoxToScreenPoint -BoundingBox $bbox
                $targetX = $screenPt.X
                $targetY = $screenPt.Y
                
                Write-DebugLog -Message "[$func] 情報: JS入力失敗。座標(X:$targetX, Y:$targetY)を物理クリック後、キーボード送信を試行します。" -Level Info
                
                if ((Get-Command -Name "Invoke-DesktopCoordinateClick" -ErrorAction SilentlyContinue) -and 
                    (Get-Command -Name "Invoke-DesktopSendKeys" -ErrorAction SilentlyContinue)) {
                    
                    # 1. 物理クリックでカーソルを合わせる。
                    Invoke-DesktopCoordinateClick -X $targetX -Y $targetY
                    
                    # 2. 全選択(Ctrl+A) -> 削除(Delete) -> 値を送信する。
                    Invoke-DesktopSendKeys -Keys "^a" -WaitMs 100 | Out-Null
                    Invoke-DesktopSendKeys -Keys "{DELETE}" -WaitMs 100 | Out-Null

                    # SendKeys特有の特殊文字をブレース {} でエスケープして純粋なテキストとして送信する。
                    $escapedValue = $Value -replace '([+^%~()\[\]{}])', '{$1}'
                    Invoke-DesktopSendKeys -Keys $escapedValue -WaitMs 100 | Out-Null

                    return "[FALLBACK_PHYSICAL] Input Completed: $Selector"
                }
            }
        } catch {}
    }

    if (-not $res) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "入力に失敗しました" -Details $Selector)
    }
    return "JS Input Completed: $Selector"
}

# --- ドロップダウンリストの指定値選択 ---
# ドロップダウン（select）の指定値を選択し、change イベントを発火する。
function Select-WebDropdown {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [Parameter(Mandatory=$true)][string]$Value,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    # Base64エンコードを利用してサニタイズする。
    $selB64 = ConvertTo-JsSafeBase64 -Text $Selector
    $valB64 = ConvertTo-JsSafeBase64 -Text $Value

    # 要素出現を待機する。
    Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null

    # value設定およびchangeイベントの発火： 値を直接設定し、変更をフレームワークに認識させる。
    $js = @"
        const selector = decodeURIComponent(escape(atob('$selB64')));
        const val = decodeURIComponent(escape(atob('$valB64')));
        const found = utilFindInFrames(window, function(win) {
            try {
                const el = deepQuerySelector(selector, win.document);
                if (el) {
                    // Select要素に指定された値をセットする。
                    el.value = val;
                    // 選択変更をブラウザ(フレームワーク)に認識させるためchangeイベントを発火させる。
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    return true;
                }
            } catch(e) {}
            return null;
        });
        return found === true;
"@

    $res = Invoke-WebScript -Js $js
    if (-not $res) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "ドロップダウンリストの選択に失敗しました" -Details $Selector)
    }
}

# --- チェックボックスの状態設定 ---
# チェックボックスの状態（True/False）を判定し、差異があれば切り替える。
function Set-WebCheckbox {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [Parameter(Mandatory=$true)][bool]$State = $true,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    # Base64エンコードを利用してサニタイズする。
    $selB64 = ConvertTo-JsSafeBase64 -Text $Selector
    $stateStr = $State.ToString().ToLower()

    # 要素出現を待機する。
    Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null

    # 状態差分が存在する場合のみのclickおよびchangeイベント発火：
    $js = @"
        const selector = decodeURIComponent(escape(atob('$selB64')));
        const targetState = $stateStr; // true または false (JSのBooleanとして直接評価される)
        const found = utilFindInFrames(window, function(win) {
            try {
                const el = deepQuerySelector(selector, win.document);
                if (el) {
                    // 現在のチェック状態が、目標とする状態(targetState)と異なる場合のみ処理を行う。
                    if (el.checked !== targetState) {
                        // 状態を反転させるためクリックを実行する。
                        el.click();
                        // 確実に目標状態となるようプロパティを直接書き換える。
                        el.checked = targetState;
                        // 変更を認識させるためchangeイベントを発火させる。
                        el.dispatchEvent(new Event('change', { bubbles: true }));
                    }
                    return true;
                }
            } catch(e) {}
            return null;
        });
        return found === true; 
"@

    $res = Invoke-WebScript -Js $js
    if (-not $res) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "チェックボックスの状態変更に失敗しました" -Details $Selector)
    }
}

# --- 指定要素のテキスト取得 ---
# 要素の innerText または value を取得する。
function Get-WebText {
    param (
        [Parameter(Mandatory=$true)][string]$Selector,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    # Base64エンコードを利用してサニタイズする。
    $selB64 = ConvertTo-JsSafeBase64 -Text $Selector

    # 要素出現を待機する。
    Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null

    # innerTextまたはvalueの返却：
    $js = @"
        const selector = decodeURIComponent(escape(atob('$selB64')));
        const text = utilFindInFrames(window, function(win) {
            try {
                const el = deepQuerySelector(selector, win.document);
                // 要素があれば String、無ければ null を返す。
                if (el) { return el.innerText || el.value || ''; }
            } catch(e) {}
            return null;
        });
        return text !== null ? text : '';
"@

    return Invoke-WebScript -Js $js
}

# --- 指定要素の存在確認 ---
# 指定要素の存在確認を行い、例外を投げず True/False の文字列を返す。
function Test-WebElement {
    param (
        [string]$Selector = "",
        [string]$XPath = "",
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    if ([string]::IsNullOrEmpty($Selector) -and [string]::IsNullOrEmpty($XPath)) {
        throw (New-EngineException -Func $func -Type "引数エラー" -Message "Selector または XPath のいずれかを指定してください。")
    }

    try {
        if (-not [string]::IsNullOrEmpty($Selector)) {
            Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec | Out-Null
        } else {
            Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec | Out-Null
        }
        return "True"
    } catch {
        # タイムアウト等、要素が見つからなかった場合は例外を握りつぶして False を返す。
        return "False"
    }
}

# --- Web要素の属性値(Attribute)を取得 ---
# 指定したWeb要素の特定の属性値を取得する。
function Get-WebAttribute {
    param (
        [string]$Selector = "",
        [string]$XPath = "",
        [string]$Attribute = "",
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name

    # 1. パラメータの必須チェック：
    if ([string]::IsNullOrEmpty($Attribute)) {
        throw (New-EngineException -Func $func -Type "引数エラー" -Message "Attribute パラメータは必須です。")
    }
    if ([string]::IsNullOrEmpty($Selector) -and [string]::IsNullOrEmpty($XPath)) {
        throw (New-EngineException -Func $func -Type "引数エラー" -Message "Selector または XPath のいずれかを指定してください。")
    }
    # Base64エンコードを利用してサニタイズする。
    $attrB64 = ConvertTo-JsSafeBase64 -Text $Attribute

    try {
        $js = ""
        # 2. Selector または XPath に応じた待機処理とJS組み立て：
        if (-not [string]::IsNullOrEmpty($Selector)) {
            # 待機関数が返す余分な "True" を $null で破棄して戻り値への混入を防ぐ。
            $null = Wait-WebElement -Selector $Selector -TimeoutSec $TimeoutSec
            
            $selB64 = ConvertTo-JsSafeBase64 -Text $Selector
            # セレクタで特定した要素の指定属性を取得して返却するJSを構築する。
            $js = "const sel = decodeURIComponent(escape(atob('$selB64'))); const attr = decodeURIComponent(escape(atob('$attrB64'))); return document.querySelector(sel).getAttribute(attr);"
        }
        elseif (-not [string]::IsNullOrEmpty($XPath)) {
            # 待機関数が返す余分な "True" を $null で破棄して戻り値への混入を防ぐ。
            $null = Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec
            
            $xpathB64 = ConvertTo-JsSafeBase64 -Text $XPath
            # XPathで特定した要素の指定属性を取得して返却するJSを構築する。
            $js = "const xpath = decodeURIComponent(escape(atob('$xpathB64'))); const attr = decodeURIComponent(escape(atob('$attrB64'))); const node = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue; return node ? node.getAttribute(attr) : null;"
        }

        # 3. JSの実行： 内部関数を利用してブラウザに命令する。
        $result = Invoke-WebView2NativeScript -Js $js
        
        # 取得結果が null の場合は空文字を返す。
        if ($result -eq "null" -or $null -eq $result) {
            return ""
        }
        return $result
    }
    catch {
        # すでに New-EngineException で生成されたフォーマット済みエラーなら、そのまま上へ投げる。
        if ($_.Exception.Message -match "^\[.*?\] \[.*?\]:") {
            throw $_
        }
        
        # 純粋なJS構文エラー等の場合のみラップして投げる。
        throw (New-EngineException -Func $func -Type "JSエラー" -Message "属性値の取得中にエラーが発生しました。" -Details $_.Exception.Message)
    }
}

# --- 現在のページURLの取得 ---
function Get-WebUrl {
    # アクティブなWebView2インスタンスから直接URLを取得する。
    $webview = Get-ActiveWebView
    if ($webview -and $webview.Source) {
        return $webview.Source.ToString()
    }
    return $null
}

# --- 現在のページタイトルの取得 ---
function Get-WebTitle {
    # JSを経由してDOMのtitleプロパティを取得する。
    try { return Invoke-WebScript -Js "return document.title;" }
    catch { return $null }
}

# --- 指定ディレクトリ・指定ファイル名への完全サイレントダウンロードを有効化 ---
# DLダイアログを抑制し、指定フォルダ・ファイル名での裏側ダウンロードを有効化する。
function Enable-SilentDownload {
    param (
        [Parameter(Mandatory = $true)][string]$DownloadDirectory,
        [string]$FileName = "" 
    )

    $func = $MyInvocation.MyCommand.Name

    try {
        if (-not (Test-Path $DownloadDirectory)) {
            # 指定された保存先ディレクトリが存在しない場合は作成する。
            New-Item -ItemType Directory -Path $DownloadDirectory -Force | Out-Null
        }

        # 保存先とファイル名をグローバル変数にセットする。（毎回の書き換えに対応）
        $global:TargetDownloadDirectory = $DownloadDirectory
        $global:TargetDownloadFileName = $FileName
        
        $webview = Get-ActiveWebView
        if ($null -eq $webview -or $null -eq $webview.CoreWebView2) {
            throw "WebView2インスタンスが取得できません"
        }

        if ($global:SilentDownloadRegistered) {
            Write-DebugLog -Message "[$func] ダウンロード先を更新: Dir=$DownloadDirectory, File=$FileName" -Level Info
            return
        }

        # new() を使わず、PowerShellの「型キャスト」を使って安全にイベントを登録する。
        $handler = [System.EventHandler[Microsoft.Web.WebView2.Core.CoreWebView2DownloadStartingEventArgs]] {
            param($sender, $e)

            # バックグラウンドスレッドを守る。
            try {
                # OS標準の保存ダイアログ(プロンプト)表示をブロックする。
                $e.Handled = $true

                # VBAからファイル名が指定されていればそれを使用し、無ければ元ファイル名を使用する。
                if ([string]::IsNullOrEmpty($global:TargetDownloadFileName)) {
                    # オリジナルのファイル名を抽出する。
                    $nameToSave = [System.IO.Path]::GetFileName($e.ResultFilePath)
                } else {
                    $nameToSave = $global:TargetDownloadFileName
                }

                # 最終的な保存先フルパスを設定する。
                $e.ResultFilePath = Join-Path $global:TargetDownloadDirectory $nameToSave
            } catch {
                Write-DebugLog -Message "[Download] 致命的エラー: サイレント保存中に例外発生 ($($_.Exception.Message))" -Level Error
            }
        }

        # 作成した安全なハンドラをセットする。
        $webview.CoreWebView2.add_DownloadStarting($handler)
        $global:SilentDownloadRegistered = $true

        Write-DebugLog -Message "[$func] サイレントダウンロードを有効化しました。" -Level Info
    } catch {
        throw (New-EngineException -Func $func -Type "初期化エラー" -Message "ダウンロード動作の設定に失敗しました" -Details $_.Exception.Message)
    }
}

# --- ダウンロード（ファイル書き込み）の完了を待機 ---
# .crdownload の消失および排他ロック解除を確認し、DL完了を待機する。
function Wait-FileDownload {
    param (
        [Parameter(Mandatory = $true)][string]$FilePath,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )
    
    $func = $MyInvocation.MyCommand.Name

    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        # WebView2のイベント（ダウンロード開始等）をブロックさせないための息継ぎ処理を行う。
        [System.Windows.Forms.Application]::DoEvents()

        # 1. ファイルが作成されているか確認：
        if (Test-Path $FilePath) {
            try {
                # 2. 一時ファイルの有無確認： Chrome/Edge特有の .crdownload ファイルが存在しないか確認する。
                $tempFile = $FilePath + ".crdownload"
                if (-not (Test-Path $tempFile)) {
                    # 3. 排他制御の解除確認： ファイルを排他モードで開けるか（書き込みロックが解除されたか）テストする。
                    $stream = [System.IO.File]::Open($FilePath, 'Open', 'Read', 'None')
                    $stream.Close()
                    
                    return $FilePath
                }
            } catch {
                # ロック中のため待機を継続する。
            }
        }
        Start-Sleep -Milliseconds 200
    }
    throw (New-EngineException -Func $func -Type "Timeout" -Message "ファイルの保存完了確認がタイムアウトしました" -Details $FilePath)
}

# --- 埋め込みPDF(embed)の絶対URLを安全に取得 ---
function Get-WebEmbedPdfUrl {
    $func = $MyInvocation.MyCommand.Name

    $js = @"
        const url = utilFindInFrames(window, function(win) {
            try {
                // embed要素を探索する。（より厳密に type 属性をチェック）
                const el = win.document.querySelector('embed[type="application/pdf"], embed[src*=".pdf"]');
                if (!el) return null;
                
                // 要素のsrc属性から元ファイルのURLを取得する。
                const rawSrc = el.getAttribute('src') || el.src;
                
                // Chromiumベースのブラウザで直リンクのPDFを開くと src="about:blank" になる仕様へ対応する。
                if (!rawSrc || rawSrc === 'about:blank') {
                    // 直接PDFを開いている場合は、現在のURLそのものを対象とする。
                    return win.location.href;
                }
                
                try {
                    // 絶対URLに変換して返す。
                    return new URL(rawSrc, win.location.href).href;
                } catch(e) {
                    return rawSrc;
                }
            } catch(e) {
                return null;
            }
        });
        return url || '';
"@

    $res = Invoke-WebScript -Js $js

    if ([string]::IsNullOrWhiteSpace($res)) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "PDFのembed要素、またはURL(src)が見つかりません")
    }
    return $res
}

# --- Fetch API を利用した裏側でのサイレントダウンロード発火 (file:// 互換版) ---
function Invoke-WebFetchDownload {
    param ([string]$TargetUrl = "")

    $func = $MyInvocation.MyCommand.Name

    # 1. 対象URLの特定： TargetUrl が指定されていない場合は現在のページURLを使用する。
    $fetchUrl = $TargetUrl

    # 2. file:// プロトコル対応： 画面遷移を防ぐため Data URI (data:application/pdf;base64,...) へ変換する。
    if ($fetchUrl.StartsWith("file://", [System.StringComparison]::OrdinalIgnoreCase)) {
        # file:/// C:/... のパスを抽出する。
        $localPath = [System.Uri]::UnescapeDataString(($fetchUrl -replace '^file:///', ''))
        $localPath = $localPath -replace '/', '\'
        
        if (Test-Path $localPath) {
            # ファイルをバイナリとして読み込み、Base64エンコードしてData URI形式のURLを再構築する。
            $bytes = [System.IO.File]::ReadAllBytes($localPath)
            $base64 = [System.Convert]::ToBase64String($bytes)
            $fetchUrl = "data:application/pdf;base64,$base64"
        }
    }

    # Base64エンコードを利用してサニタイズする。
    $targetUrlB64 = ConvertTo-JsSafeBase64 -Text $fetchUrl

    # 3. Fetch API コアロジックの実行：
    <#
      Fetch APIを利用して対象URLのデータをバックグラウンドで取得し、
      メモリ上に生成したBlobから仮想的なアンカータグをクリックさせてダウンロードを強制発火させる。
    #>
    $js = @"
        try {
            let target = decodeURIComponent(escape(atob('$targetUrlB64')));
            if (!target) target = window.location.href;

            fetch(target)
                // 取得したデータをバイナリオブジェクト(Blob)へ変換する。
                .then(r => r.blob())
                .then(b => {
                    // Blobから一時的なオブジェクトURLを生成する。
                    const url = window.URL.createObjectURL(b);
                    // 見えないアンカータグ(a)を動的に生成し、ダウンロード属性を付与する。
                    const a = document.createElement('a');
                    a.href = url;
                    // 実際のファイル名はエンジン側で上書きする。
                    a.download = 'auto_download_temp';
                    // DOMへ追加し、強制的にクリックイベントを発火させてダウンロードを開始する。
                    document.body.appendChild(a);
                    a.click();
                    
                    // メモリ解放とDOMのお掃除を実行する。
                    setTimeout(function() {
                        // 一時URLを破棄し、アンカータグをDOMツリーから削除する。
                        window.URL.revokeObjectURL(url);
                        if (a.parentNode) a.parentNode.removeChild(a);
                    }, 1000);
                });
            return true;
        } catch(e) {
            return false;
        }
"@

    $res = Invoke-WebScript -Js $js
    
    if (-not $res) {
        throw (New-EngineException -Func $func -Type "JSエラー" -Message "Fetch APIを利用したダウンロードトリガーの発火に失敗しました" -Details $TargetUrl)
    }
    return "Fetch Download Triggered"
}
