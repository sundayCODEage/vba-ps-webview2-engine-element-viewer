# ------------------------------------------------------------------------------#●#✘ 
<#
  [XPath 専用 Web 自動化操作モジュール]： document.evaluateを利用した動的要素の特定および操作を定義する。
#>
# ------------------------------------------------------------------------------

# --- XPathの正規化 ---
# XPathの表記揺れ（改行・空白）を自動補正する（内部関数）。
function Normalize-XPath {
    param([string]$XPath)

    # normalize-space包含時はスキップする。
    if ($XPath -match "normalize-space") {
        return $XPath
    }

    # 完全一致から部分一致への変換（テキスト検索）： [text()='xxx'] から、空白を無視した contains(normalize-space(.), 'xxx') へ変換する。
    $pattern = "\[[^\]]*text\(\)\s*=\s*'([^']+)'\]"
    if ($XPath -match $pattern) {
        $value = $Matches[1]
        return ($XPath -replace $pattern, "[contains(normalize-space(.), '$value')]")
    }

    # 完全一致から部分一致への変換（ドット自身）： [.='xxx'] から、contains(normalize-space(.), 'xxx') へ変換する。
    $pattern2 = "\[[^\]]*\.\s*=\s*'([^']+)'\]"
    if ($XPath -match $pattern2) {
        $value = $Matches[1]
        return ($XPath -replace $pattern2, "[contains(normalize-space(.), '$value')]")
    }
    return $XPath
}

# --- 指定XPath要素の出現および可視化待機 ---
# XPath指定で要素がDOM上に出現し、かつ可視化されるまで待機する。デバッグ時は赤枠ハイライトを実行する。
function Wait-WebXPathElement {
    param (
        [Parameter(Mandatory = $true)][string]$XPath,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    # 渡されたXPathを安全な構文へ正規化する。
    $xpathNorm = Normalize-XPath -XPath $XPath
    # Base64エンコードを利用してサニタイズする。
    $xpathB64 = ConvertTo-JsSafeBase64 -Text $xpathNorm

    # Wait-Conditionに渡すための条件判定処理： iframe内部を透過探索し、要素が可視化されるまで待機する。
    $condition = {
        try {
            $js = @"
                // Base64化されたXPathをデコードして復元する。
                const xpath = decodeURIComponent(escape(atob('$xpathB64')));
                // 共通関数を利用し、iframe内も含めて要素を透過探索する。
                const found = utilFindInFrames(window, function(win) {
                    try {
                        // evaluateを利用し、XPathに合致する最初の要素(FIRST_ORDERED_NODE_TYPE)を取得する。
                        const el = win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                        if (el) {
                            // 要素が存在する場合、スタイルとサイズを取得して可視状態を判定する。
                            const style = win.getComputedStyle(el);
                            const rect  = el.getBoundingClientRect();
                            const visible = (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && rect.width > 0 && rect.height > 0);
                            // 可視状態であれば 'visible' を返す。
                            return visible ? 'visible' : null;
                        }
                    } catch(e) {}
                    return null;
                });
                return found || 'not_found';
"@

            $res = Invoke-WebScript -Js $js

            # 要素が可視状態として検出された場合の処理を実行する。
            if ($res -eq "visible") {
                # 任意ハイライト処理を実行する。
                if ($global:CONFIG.EnableHighlight) {
                    $jsHighlight = @"
                        // Base64化されたXPathをデコードして復元する。
                        const xpath = decodeURIComponent(escape(atob('$xpathB64')));
                        // 共通関数を利用し、iframe内も含めて要素を透過探索する。
                        const el = utilFindInFrames(window, function(win) {
                            try { return win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue || null; }
                            catch(e) { return null; }
                        });

                        if (el) {
                            // 元の枠線スタイルを退避し、一時的に赤枠を適用する。
                            const oldOutline = el.style.outline;
                            el.style.outline = '2px solid red';
                            setTimeout(() => { el.style.outline = oldOutline; }, 800);
                        }
"@

                    Invoke-WebScript -Js $jsHighlight | Out-Null
                }
                return $true
            }
        } catch {}
        return $false
    }
    $errMsg = "[$func] タイムアウト: XPath要素が見つからない ($XPath)"
    Wait-Condition -ConditionBlock $condition -TimeoutSec $TimeoutSec -TimeoutMessage $errMsg | Out-Null

    return $true
}

# --- 指定XPath要素の消滅または非表示待機 ---
# XPath要素の非表示、またはDOMからの消滅を待機する。
function Wait-WebXPathElementDisappear {
    param (
        [Parameter(Mandatory = $true)][string]$XPath,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    # 渡されたXPathを安全な構文へ正規化する。
    $xpathNorm = Normalize-XPath -XPath $XPath
    # Base64エンコードを利用してサニタイズする。
    $xpathB64 = ConvertTo-JsSafeBase64 -Text $xpathNorm

    $js = @"
        // Base64化されたXPathをデコードして復元する。
        const xpath = decodeURIComponent(escape(atob('$xpathB64')));
        // 共通関数を利用し、iframe内も含めて要素を透過探索する。
        const found = utilFindInFrames(window, function(win) {
            try {
                // evaluateを利用し、XPathに合致する最初の要素(FIRST_ORDERED_NODE_TYPE)を取得する。
                const el = win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue;
                if (el) {
                    const style = win.getComputedStyle(el);
                    const rect = el.getBoundingClientRect();
                    // display, visibility, opacity, またはサイズのいずれかで非表示になれば false 扱いとする。
                    const isVisible = (style.display !== 'none' && style.visibility !== 'hidden' && style.opacity !== '0' && rect.width > 0 && rect.height > 0);
                    if (isVisible) return true; // まだ表示されている
                }
            } catch(e) {}
            return null;
        });
        return found === true;
"@

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $isGone = $false

    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        # DOM遷移中等の通信切断を避けるため、CDPではなくネイティブAPIを用いて判定を繰り返す。
        $exists = Invoke-WebView2NativeScript -Js $js -Retries 1
        # 戻り値が false (非表示/消滅) となった場合はループを抜ける。
        if ($exists -eq $false -or $exists -match "false") {
            $isGone = $true
            break
        }
        Start-Sleep -Milliseconds 500
    }

    if (-not $isGone) {
        throw (New-EngineException -Func $func -Type "Timeout" -Message "指定された時間が経過しても要素が消滅・非表示になりませんでした" -Details "${TimeoutSec}秒経過 ($XPath)")
    }
}

# --- 指定XPath要素のクリック実行 ---
# XPath要素に対し、一連のマウスイベントをエミュレートしクリックを実行する。失敗時は物理操作へフォールバックする。
function Invoke-WebXPathClick {
    param (
        [Parameter(Mandatory = $true)][string]$XPath,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec,
        [string]$OuterHtml = "",
        [string]$BoundingBox = ""
    )

    $func = $MyInvocation.MyCommand.Name
    # 渡されたXPathを安全な構文へ正規化する。
    $xpathNorm = Normalize-XPath -XPath $XPath
    # Base64エンコードを利用してサニタイズする。
    $xpathB64 = ConvertTo-JsSafeBase64 -Text $xpathNorm

    # 要素が出現し操作可能になるまで待機する。
    Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec | Out-Null

    $js = @"
        // Base64化されたXPathをデコードして復元する。
        const xpath = decodeURIComponent(escape(atob('$xpathB64')));
        // 共通関数を利用し、iframe内も含めて要素を透過探索する。
        const el = utilFindInFrames(window, function(win) {
            try {
                return win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue || null;
            } catch(e) { return null; }
        });

        if (el) {
            // 要素を画面の中央へスクロールさせる。
            el.scrollIntoView({block:'center', inline:'center'});

            // 人間の操作を模倣するため、一連のマウスイベントを強制的に発火(ディスパッチ)させる。
            el.dispatchEvent(new MouseEvent('mouseover', {bubbles:true}));
            el.dispatchEvent(new MouseEvent('mousemove', {bubbles:true}));
            el.dispatchEvent(new MouseEvent('mousedown', {bubbles:true}));
            el.dispatchEvent(new MouseEvent('mouseup', {bubbles:true}));

            // フォーカスを当ててからクリックメソッドを呼び出す。
            el.focus();
            el.click();
            return true;
        }
        return false;
"@

    $res = Invoke-WebScript -Js $js
    
    if (-not $res) { Write-DebugLog -Message "[$func] 警告: JSによるXPath要素のクリックが失敗しました。" -Level Warn }

    # JSクリック失敗時の物理クリックフォールバック： ピッカーからの座標(BoundingBox)指定がある場合はOS物理クリックへ移行する。
    if (-not [string]::IsNullOrWhiteSpace($BoundingBox)) {
        try {
            # JSON形式のBoundingBoxをPowerShellオブジェクトに変換する。
            $bbox = $BoundingBox | ConvertFrom-Json
             if ($bbox.width -gt 0 -and $bbox.height -gt 0) {
                # 論理座標をOSの絶対スクリーン座標へ変換する。
                $screenPt = Convert-BoundingBoxToScreenPoint -BoundingBox $bbox
                $targetX = $screenPt.X
                $targetY = $screenPt.Y

                Write-DebugLog -Message "[$func] 情報: BoundingBoxを使用し、座標(X:$targetX, Y:$targetY)へ物理クリックを試行します。" -Level Info
                # DesktopUIAモジュール(Win32API)の物理クリック関数がロードされているか確認する。
                if (Get-Command -Name "Invoke-DesktopCoordinateClick" -ErrorAction SilentlyContinue) {
                    Invoke-DesktopCoordinateClick -X $targetX -Y $targetY
                    return $true
                } else {
                    Write-DebugLog -Message "[$func] 警告: 物理クリック関数が見つかりません。Lib-DesktopUIAとの連携を確認してください。" -Level Warn
                }
            }
        } catch {
            Write-DebugLog -Message "[$func] エラー: BoundingBoxの解析または物理クリック中に例外が発生しました。" -Level Error
        }
    }

    if (-not $res) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "XPath要素のクリック実行に失敗しました" -Details $XPath)
    }
}

# --- 指定XPath要素へのテキスト入力 ---
# XPath要素へフォーカスし、テキスト入力と各種イベント発火を実行する。
function Set-WebXPathTextInput {
    param (
        [Parameter(Mandatory = $true)][string]$XPath,
        [Parameter(Mandatory = $true)][string]$Value,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    # 渡されたXPathを安全な構文へ正規化する。
    $xpathNorm = Normalize-XPath -XPath $XPath
    # Base64エンコードを利用してサニタイズする。
    $xpathB64 = ConvertTo-JsSafeBase64 -Text $xpathNorm
    $valB64 = ConvertTo-JsSafeBase64 -Text $Value

    # 要素が出現し操作可能になるまで待機する。
    Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec | Out-Null

    $js = @"
        // Base64化されたXPathをデコードして復元する。
        const xpath = decodeURIComponent(escape(atob('$xpathB64')));
        const val = decodeURIComponent(escape(atob('$valB64')));
        // 共通関数を利用し、iframe内も含めて要素を透過探索する。
        const el = utilFindInFrames(window, function(win) {
            try {
                return win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue || null;
            } catch(e) { return null; }
        });

        if (el) {
            // 対象要素にフォーカスを当てる。
            el.focus();
            // DOMのプロパティとして値を直接書き込む。
            el.value = val;
            // ReactやVueなどのモダンフレームワークに変更を認識させるため、inputとchangeイベントを強制発火させる。
            el.dispatchEvent(new Event('input',  { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            return true;
        }
        return false;
"@

    $res = Invoke-WebScript -Js $js
    if (-not $res) {
        throw (New-EngineException -Func $func -Type "未発見" -Message "XPath要素へのテキスト入力に失敗しました" -Details $XPath)
    }
}

# --- 指定XPath要素のテキスト取得 ---
# XPath要素のタグを判別し、適切なテキスト（value または innerText）を取得する。
function Get-WebXPathText {
    param (
        [Parameter(Mandatory = $true)][string]$XPath,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec
    )

    $func = $MyInvocation.MyCommand.Name
    # 渡されたXPathを安全な構文へ正規化する。
    $xpathNorm = Normalize-XPath -XPath $XPath
    # Base64エンコードを利用してサニタイズする。
    $xpathB64 = ConvertTo-JsSafeBase64 -Text $xpathNorm

    # 要素が出現するまで待機する。
    Wait-WebXPathElement -XPath $XPath -TimeoutSec $TimeoutSec | Out-Null

    $js = @"
        // Base64化されたXPathをデコードして復元する。
        const xpath = decodeURIComponent(escape(atob('$xpathB64')));
        // 共通関数を利用し、iframe内も含めて要素を透過探索する。
        const el = utilFindInFrames(window, function(win) {
            try {
                return win.document.evaluate(xpath, win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue || null;
            } catch(e) { return null; }
        });

        // タグの種類に応じて取得するプロパティを切り替える。
        if (!el) throw new Error("XPath element not found: " + xpath);

        const tag = el.tagName ? el.tagName.toLowerCase() : '';
        if (tag === 'input' || tag === 'textarea' || tag === 'select') {
            // 入力フォーム系の要素であれば value プロパティを返す。
            return el.value;
        } else {
            // それ以外の要素であれば表示されているテキスト(innerText/textContent)を返す。
            return el.innerText || el.textContent;
        }
"@

    return Invoke-WebScript -Js $js
}
