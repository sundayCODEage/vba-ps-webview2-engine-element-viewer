# ------------------------------------------------------------------------------#●#✘ 
<#
  デバッグおよび証跡ファイルエクスポートモジュール
  画面状態のスナップショット保存およびDOM/iframe/Window情報の解析等
#>
# ------------------------------------------------------------------------------

# ==============================================================================
# [1. 画面状態・スナップショットの保存]： HTMLやスクリーンショットをログディレクトリへ出力する。

# --- 画面全体（多段iframe含む）のHTML保存 ---
# クロスオリジンを考慮し、全iframeを含むHTMLスナップショットを保存する。
function Export-WebHtml {
    param ([string]$Prefix = "WebHtml")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $filePath  = Join-Path $global:LogDir "${Prefix}_${timestamp}.html"

    try {
        $js = @"
        // 指定されたウィンドウ(win)とその子フレームを再帰的に巡回してHTML文字列を抽出する関数
        function dump(win, depth) {
            let url = "UNKNOWN";
            // CORS対策： クロスドメインのiframeにアクセスすると例外が発生するため保護する。
            try { url = win.location.href; } catch(e) { url = "CROSS_ORIGIN_DENIED"; }
            
            // 取得したURLをHTMLの先頭にコメントとして明記する（デバッグ時の追跡用）。
            let html = "\n<!-- [FRAME URL: " + url + "] -->\n";
            
            try {
                // HTMLタグを含むドキュメント全体の文字列を取得して追加する。
                html += win.document.documentElement.outerHTML + "\n";
            } catch(e) {
                // アクセス拒否時はコメントのみを出力する。
                html += "<!-- DOM Access Denied -->\n";
            }

            // 子フレームが存在する場合はループ処理を行う。
            for (let i = 0; i < win.frames.length; i++) {
                // 子フレームを引数として再帰呼び出しを行い、結果を結合する。
                try { html += dump(win.frames[i], depth + 1); } catch(e) {}
            }
            return html;
        }
        return dump(window, 0);
"@

        $html = Invoke-WebScript -Js $js
        # UTF-8エンコーディングでHTMLファイルとして保存する。
        $html | Out-File -FilePath $filePath -Encoding UTF8 -Force
        return "[OK] Export-WebHtml: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" `
            -Message "HTMLファイルの保存に失敗しました" -Details $_.Exception.Message)
    }
}

# --- 画面スクリーンショット（CDP/WebView2）のPNG保存 ---
# CDP、またはネイティブAPIへフォールバックして画面のPNGスクショを保存する。
function Export-WebScreenshot {
    param ([string]$Prefix = "WebScreenshot")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $guid      = [guid]::NewGuid().ToString()
    $fileName  = "${Prefix}_${timestamp}_${guid}.png"
    $filePath  = Join-Path $global:LogDir $fileName

    try {
        $useNative = $false
        # CDPモードでのスクリーンショット撮影を試行する。
        if ($global:CdpPort -gt 0) {
            try {
                # リトライとタイムアウトを短めに設定してCDP撮影を試みる。
                $res = Invoke-CdpCommand -Method "Page.captureScreenshot" -Params @{ format = "png" } -MaxRetries 1 -TimeoutSecPerTry 3
                # Base64文字列をバイナリ配列に変換する。
                $bytes = [Convert]::FromBase64String($res.result.data)
                # バイナリデータをPNGファイルとしてディスクへ書き込む。
                [IO.File]::WriteAllBytes($filePath, $bytes)
            } catch {
                Write-DebugLog -Message "[$func] 警告: CDP撮影失敗。ネイティブAPI(CapturePreview)へフォールバックします ($($_.Exception.Message))" -Level Warning
                $useNative = $true
            }
        } else {
            $useNative = $true
        }

        # CDP無効時、またはCDP失敗時はネイティブAPIでの撮影を実行する。
        if ($useNative) {
            # 画像データを受け取るためのメモリストリームを初期化する。
            $stream = [IO.MemoryStream]::new()
            $webview = Get-ActiveWebView
 
            # WebView2のネイティブAPIを非同期で呼び出し、ストリームへ画像データを書き込ませる。            
            $task = $webview.CoreWebView2.CapturePreviewAsync(
                [Microsoft.Web.WebView2.Core.CoreWebView2CapturePreviewImageFormat]::Png,
                $stream
            )

            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $timeoutSec = 10  # タイムアウトまでの許容時間

            # 非同期タスクが完了するまでUIをフリーズさせずに待機する。           
            while (-not $task.IsCompleted) {
                if ($sw.Elapsed.TotalSeconds -gt $timeoutSec) {
                    throw (New-EngineException -Func $func -Type "Timeout" `
                        -Message "ネイティブスクショ取得がタイムアウトしました" -Details "${timeoutSec}秒経過")
                }
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 10
            }

            if ($task.IsFaulted) {
                throw (New-EngineException -Func $func -Type "ネイティブエラー" `
                    -Message "CapturePreviewAsyncAPI失敗" -Details $task.Exception.InnerException.Message)
            }

            # ファイル書き込み中の例外発生時も確実にメモリストリームを解放する。
            $fileStream = [IO.File]::Create($filePath)
            try {
                # ストリームの読み取り位置を先頭に戻し、ファイルストリームへ一括コピーする。
                $stream.Seek(0, [IO.SeekOrigin]::Begin) | Out-Null
                $stream.CopyTo($fileStream)
            } finally {
                if ($null -ne $fileStream) { $fileStream.Dispose() }
                if ($null -ne $stream) { $stream.Dispose() }
            }
        }
        return "[OK] Export-WebScreenshot: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" `
            -Message "スクリーンショットの保存処理中に例外が発生しました" -Details $_.Exception.Message)
    }
}

# ==============================================================================
# [2. DOM・要素・フレーム構造の解析とエクスポート]： テーブルや操作可能要素をCSVとして保存する。

# --- 指定テーブルのCSV保存およびVBA連携用データの返却 ---
# テーブル要素を解析し、VBA取込用の配列文字列を含むCSVを生成する。
function Export-WebTableToCsv {
    param (
        [Parameter(Mandatory = $true)][string]$Selector,
        [string]$FileName = "WebTable.csv"
    )

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    # Base64エンコードを利用してサニタイズする。
    $selB64 = ConvertTo-JsSafeBase64 -Text $Selector

    try {
        $js = @"
            $global:ENGINE_JS_UTILS
            const sel = decodeURIComponent(escape(atob('$selB64')));
            try {
                // 共通関数を利用し、iframe内も含めて対象のtable要素を探索する。
                const table = utilFindInFrames(window, function(win) {
                    try { return win.document.querySelector(sel) || null; }
                    catch(e) { return null; }
                });

                if (!table) return 'NOT_FOUND';

                // テーブルの行(TR)要素のコレクションを配列へ変換する。
                const rows = Array.from(table.rows);
                if (rows.length === 0) return 'EMPTY_TABLE';

                // CSV出力用の文字列を生成する（人間が読む用）
                const csvData = rows.map(row => {
                    return Array.from(row.cells).map(cell => {
                        // 改行文字をスペースに置換し、ダブルクォートでエスケープする。
                        const text = (cell.innerText || '').trim().replace(/\r?\n|\r/g, ' ');
                        return '"' + text.replace(/"/g, '""') + '"';
                    }).join(',');
                }).join('\r\n');

                // VBA連携用のデータ文字列を生成する（データ抽出用）
                const vbaData = rows.map(row => {
                    return Array.from(row.cells).map(cell => {
                        // 1. テキストがあればそれを返す。
                        const text = (cell.innerText || '').trim().replace(/\r?\n|\r/g, ' ');
                        if (text !== '') return text;

                        // 2. テキストが空なら画像ファイル名を返す（汎用処理）
                        const img = cell.querySelector('img');
                        if (img && img.src) {
                            const parts = img.src.split('/');
                            return parts[parts.length - 1]; // 例: "IP10B030.png"
                        }
                        return '';
                    // VBA側でのSplit解析 (セル間を<T>、行間を<R>という独自デリミタで結合する)
                    }).join('<T>');
                }).join('<R>');

                return { csv: csvData, vba: vbaData };
            } catch(e) {
                return 'JS_EXCEPTION: ' + (e && e.message ? e.message : String(e));
            }
"@

        $result = Invoke-WebScript -Js $js

        if ($result -eq 'NOT_FOUND')    { throw (New-EngineException -Func $func -Type "未発見" -Message "指定されたテーブルが見つかりません" -Details $Selector) }
        if ($result -eq 'EMPTY_TABLE')  { throw (New-EngineException -Func $func -Type "未発見" -Message "テーブル内に行データが存在しません" -Details $Selector) }
        if ($result -match '^JS_EXCEPTION:') { throw (New-EngineException -Func $func -Type "JSエラー" -Message "テーブル解析中にJavaScript例外が発生しました" -Details $result) }

        if ($FileName -ne "") {
            $savePath = Join-Path $global:LogDir $FileName
            $result.csv | Out-File -FilePath $savePath -Encoding UTF8 -Force
        }

        # 注意： 他のExport系関数とは異なり、VBA側でメモリ上で直接データとして活用(配列化等)することを想定する（生データを返却する）。
        return $result.vba

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" `
            -Message "テーブルのCSV保存処理中に例外が発生しました" -Details $_.Exception.Message)
    }
}

# --- 操作可能要素（input/button/a等）の抽出およびCSV保存 ---
# 画面内の操作可能要素の属性を総ざらいしてCSV化する。
function Export-WebElementsToCsv {
    param (
        [string]$FileName = "WebElements.csv",
        [int]$Limit = 1000
    )

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    $js = @"
        // ドキュメント内の全要素を巡回し、操作可能な要素を抽出して配列へ格納する再帰関数
        function dump(win, arr) {
            const els = win.document.querySelectorAll('*');
            for (const el of els) {
                const tag = el.tagName.toLowerCase();
                // RPAの操作対象になりやすい主要タグ、またはIDを持つ要素を抽出対象とする。
                if (['a','button','input','select','textarea'].includes(tag) || el.id) {
                    const style = win.getComputedStyle(el);
                    // styleがnullになるエッジケースを考慮し、デフォルト値を空文字にする。
                    const displayVal = style ? style.display : 'none';
                    const visibilityVal = style ? style.visibility : 'hidden';
 
                    // 解析に有用な属性値をオブジェクトとして配列に追加する。                    
                    arr.push({
                        TagName: tag,
                        Type: el.type || '',
                        Name: el.name || '',
                        ID: el.id || '',
                        Value: (el.value || '').substring(0,50),
                        Text: (el.innerText || '').substring(0,50),
                        OuterHTML: (el.outerHTML || '').substring(0,200),
                        // CSSスタイルから要素が画面に表示されているかを判定する。
                        IsDisplayed: (displayVal !== 'none' && visibilityVal !== 'hidden')
                    });
                }
            }
            // 子フレームに対しても再帰的に抽出処理を行う。
            for (let i = 0; i < win.frames.length; i++) {
                try { dump(win.frames[i], arr); } catch(e) {}
            }
            return arr;
        }
        return dump(window, []);
"@

    try {
        $elements = Invoke-WebScript -Js $js

        if ($null -eq $elements -or $elements.Count -eq 0) {
            Write-DebugLog -Message "[$func] 情報: 要素なし" -Level Info
            return
        }

        $results = @()
        $idx = 1

        foreach ($el in $elements) {
            $results += [PSCustomObject]@{
                No          = $idx
                TagName     = $el.TagName
                Type        = $el.Type
                Name        = $el.Name
                ID          = $el.ID
                Value       = $el.Value
                Text        = $el.Text
                OuterHTML   = $el.OuterHTML
                IsDisplayed = $el.IsDisplayed
            }
            $idx++
            if ($idx -gt $Limit) { break }
        }

        $filePath = Join-Path $global:LogDir $FileName
        # パイプラインを利用してデータを蓄積し、CSVファイルとして出力する。
        $results | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Force
        return "[OK] Export-WebElementsToCsv: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" `
            -Message "Web要素のCSVエクスポートに失敗しました" -Details $_.Exception.Message)
    }
}

# --- iframe/frame階層構造のツリー形式CSV保存 ---
# 多段iframeのネスト構造をツリー形式で解析しCSV化する。
function Export-WebFrameTreeToCsv {
    param ([string]$FileName = "WebFrameTree.csv")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    $js = @"
        // iframeの親子関係を再帰的に解析し、ツリー構造のオブジェクト配列として返す関数
        function getFrameTree(win, depth, path) {
            const frames = win.document.querySelectorAll('iframe, frame');
            let tree = [];

            for (let i = 0; i < frames.length; i++) {
                const f = frames[i];
                // 階層パスをドット区切りで生成する（例: "0.1.0"）。
                const currentPath = path === '' ? i.toString() : path + '.' + i;

                const info = {
                    Depth: depth,
                    Path: currentPath,
                    Index: i,
                    Tag: f.tagName.toLowerCase(),
                    Id: f.id || '',
                    Name: f.name || '',
                    Src: f.src || ''
                };

                try {
                    const hasAccess = f.contentWindow && f.contentWindow.document;
                    if (hasAccess) {
                        // アクセス可能なフレームであれば再帰呼び出しを行い、子フレームの情報を結合する。
                        const children = getFrameTree(f.contentWindow, depth + 1, currentPath);
                        tree.push(info);
                        tree = tree.concat(children);
                    } else {
                        // クロスドメイン制約(CORS)によりアクセス不能なフレームであることを記録する。
                        info.Note = "Cross-Origin (No Access)";
                        tree.push(info);
                    }
                } catch(e) {
                    info.Note = "Access Denied";
                    tree.push(info);
                }
            }
            return tree;
        }
        // 最終結果をJSON文字列化してPowerShellへ返却する。
        return JSON.stringify(getFrameTree(window, 0, ''));
"@

    try {
        $json = Invoke-WebScript -Js $js

        if ([string]::IsNullOrWhiteSpace($json) -or $json -eq "[]") {
            Write-DebugLog -Message "[$func] 情報: iframeなし" -Level Info
            return
        }

        try {
            $frames = $json | ConvertFrom-Json -ErrorAction Stop
        } catch {
            throw (New-EngineException -Func $func -Type "内部エラー" -Message "フレームツリーJSONのデコードに失敗しました" -Details $_.Exception.Message)
        }

        $results = foreach ($f in $frames) {
            $indent = "  " * $f.Depth
            $prefix = if ($f.Depth -eq 0) { "■" } else { "└─" }

            [PSCustomObject]@{
                VisualHierarchy = "$indent$prefix $($f.Tag)"
                IndexPath       = $f.Path
                Id              = $f.Id
                Name            = $f.Name
                Depth           = $f.Depth
                Tag             = $f.Tag
                Src             = $f.Src
                Note            = $f.Note
            }
        }

        $filePath = Join-Path $global:LogDir $FileName
        # パイプラインを利用してデータを蓄積し、CSVファイルとして出力する。
        $results | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Force
        return "[OK] Export-WebFrameTreeToCsv: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" `
            -Message "フレームツリーの解析またはCSV保存に失敗しました" -Details $_.Exception.Message)
    }
}

# ==============================================================================
# [3. OSウィンドウレベルの解析・保存]： アプリのウィンドウ全体やプロセス一覧を記録する。

# --- RPAブラウザウィンドウ全体（枠含む）のPNG保存 ---
# Win32 API等を使用し、ブラウザの枠を含むウィンドウ全体のスクショを保存する。
function Export-WindowScreenshot {
    param ([string]$Prefix = "WindowScreenshot")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    try {
        Add-Type -AssemblyName System.Drawing

        if ($null -eq $global:BrowserForm) {
            throw (New-EngineException -Func $func -Type "未発見" `
                -Message "撮影対象のRPAブラウザウィンドウが存在しません")
        }

        # RPAブラウザフォーム(WinForms)の論理的な境界短形を取得する。
        $bounds = $global:BrowserForm.Bounds

        # --- DPIスケーリング(125%など)の補正計算 ---
        # OSレベルのGraphicsオブジェクトを生成し、X軸のDPIを取得する。
        $tmpGraphics = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
        $dpiX = $tmpGraphics.DpiX
        $tmpGraphics.Dispose()
        # 標準DPI(96)で除算し、スケーリング倍率を算出する。
        $scale = $dpiX / 96.0  # 100%時は1.0、125%時は1.25になる。

        # --- 物理ピクセル座標・サイズへの変換 ---
        # 論理座標とサイズにスケーリング倍率を乗算し、スクリーンショット用の物理ピクセル値を確定する。
        $phyX = [int][Math]::Round($bounds.X * $scale)
        $phyY = [int][Math]::Round($bounds.Y * $scale)
        $phyW = [int][Math]::Round($bounds.Width * $scale)
        $phyH = [int][Math]::Round($bounds.Height * $scale)

        # 補正後の物理サイズでBitmapを作成する。
        $bmp      = New-Object System.Drawing.Bitmap $phyW, $phyH
        $graphics = [System.Drawing.Graphics]::FromImage($bmp)

        # 物理ピクセルを指定してOSレベルの画面キャプチャを実行する。
        # 画面上の指定座標(ウィンドウ位置)から、Bitmapへ画像を直接コピーする。
        $graphics.CopyFromScreen(
            $phyX, $phyY,
            0, 0,
            $bmp.Size,
            [System.Drawing.CopyPixelOperation]::SourceCopy
        )

        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $guid      = [guid]::NewGuid().ToString()
        $fileName  = "${Prefix}_${timestamp}_${guid}.png"
        $filePath = Join-Path $global:LogDir $fileName

        $bmp.Save($filePath, [System.Drawing.Imaging.ImageFormat]::Png)

        $graphics.Dispose()
        $bmp.Dispose()

        return "[OK] Export-WindowScreenshot: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "内部エラー" `
            -Message "ウィンドウ全体スクリーンショットの撮影または保存に失敗しました" -Details $_.Exception.Message)
    }
}

# --- 実行中ウィンドウ一覧のCSV保存 ---
# OS上で起動している全プロセスのハンドルとタイトル一覧をCSV出力する。
function Export-WindowHierarchyToCsv {
    param ([string]$FileName = "WindowHierarchy.csv")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    try {
        # OS上で稼働中の全プロセスのうち、メインウィンドウハンドルを持つ(画面が存在する)プロセスを抽出する。
        $processes = Get-Process | Where-Object { $_.MainWindowHandle -ne 0 }
        $results = foreach ($p in $processes) {
            [PSCustomObject]@{
                ProcessName = $p.ProcessName
                Handle      = $p.MainWindowHandle
                Title       = $p.MainWindowTitle
                ID          = $p.Id
            }
        }

        $filePath = Join-Path $global:LogDir $FileName
        # パイプラインを利用してデータを蓄積し、CSVファイルとして出力する。
        $results | Export-Csv -Path $filePath -NoTypeInformation -Encoding UTF8 -Force
        return "[OK] Export-WindowHierarchyToCsv: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "内部エラー" `
            -Message "OSウィンドウ一覧の取得またはCSV保存に失敗しました" -Details $_.Exception.Message)
    }
}

# ==============================================================================
# [4. 汎用デバッグユーティリティ]： 任意のデバッグ文字列をテキストに保存するなど、汎用的な補助機能を提供する。

# --- 任意のデバッグ文字列のテキストファイル追記 ---
# 任意の文字列をデバッグ用テキストファイルへ追記保存する。
function Write-DebugTextFile {
    param (
        [Parameter(Mandatory = $true)][string]$Text,
        [string]$FileName = "DebugMemo.txt"
    )

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }
    $filePath = Join-Path $global:LogDir $FileName

    try {
        # 指定されたテキストを末尾に追記(Append)モードでファイルへ書き込む。
        $Text | Out-File -FilePath $filePath -Append -Encoding UTF8 -Force
        return "[OK] Write-DebugTextFile: $filePath"

    } catch {
        throw (New-EngineException -Func $func -Type "ファイルエラー" `
            -Message "デバッグメモのファイル書き込みに失敗しました" -Details $_.Exception.Message)
    }
}

# ==============================================================================

# --- 画面上のDOMツリーとスタイルをCDP経由で一括取得 ---
# 画面上のDOMツリーとスタイルをCDP経由で一括取得し、JSONとして保存する。
function Export-WebDomSnapshot {
    param([string]$Prefix = "DOMSnapshot")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filePath  = Join-Path $global:LogDir "${Prefix}_${timestamp}.json"

        # デザイン系を排除し、表示状態の判定に必要なスタイル要素のみを厳選する。
        $params = @{
            computedStyles = @("display", "visibility", "width", "height")
        }
        # CDPの DOMSnapshot.captureSnapshot メソッドを発行する。
        $cdpResponse = Invoke-CdpCommand -Method "DOMSnapshot.captureSnapshot" -Params $params

        # 応答から .result 本体を取得する。
        if (-not $cdpResponse -or -not $cdpResponse.result) {
            throw (New-EngineException -Func $func -Type "CDPエラー" `
                -Message "DOM Snapshot の取得結果が空です" -Details "Invoke-CdpCommand の戻り値が null / 空でした")
        }

        # Convert-DomSnapshotFlat を適用して人間が読める形式へ変換： CDP特有の多次元・圧縮配列データを、解析しやすいフラットなオブジェクトリストへ展開する。
        $flatNodes = Convert-DomSnapshotFlat -domResult $cdpResponse.result
        # JSON化して保存する (Depthを適正化し、見やすくインデント整形。意図的に -Compress を外す)。
        $json = $flatNodes | ConvertTo-Json -Depth 5
        
        try {
            Set-Content -Path $filePath -Value $json -Encoding UTF8
        }
        catch {
            throw (New-EngineException -Func $func -Type "ファイルエラー" `
                -Message "DOM Snapshot の保存に失敗しました" -Details $_.Exception.Message)
        }
        return "[OK] Export-WebDomSnapshot: $filePath"

    } catch {
        # すでに New-EngineException で生成されたエラーならそのままスローする。
        if ($_.Exception.Message -match "^\[.*?\] \[.*?\]:") { throw $_ }
        
        throw (New-EngineException -Func $func -Type "内部エラー" `
            -Message "DOM Snapshot の取得処理中に例外が発生しました" -Details $_.Exception.Message)
    }
}

# --- CDPのフラットなDOM配列をオブジェクトリストに展開 ---
# CDP固有の圧縮された1次元配列を、解析しやすいフラットなオブジェクトリストに展開する。(内部関数)
<#
  CDPの仕様（圧縮された1次元配列）と、PowerShell 5.1の処理性能の制約を克服するため、辞書データとノード・属性・スタイルを高速に再結合して扱いやすい構造へ変換する。
#>
function Convert-DomSnapshotFlat {
    param(
        [Parameter(Mandatory=$true)]
        [object]$DomResult,

        # リクエスト時に指定したスタイルの順番を配列で定義する。
        [string[]]$StyleNames = @("display", "visibility", "width", "height")
    )

    $list = New-Object System.Collections.Generic.List[object]
    # CDPから返却される「文字列の実体」がすべて格納された辞書配列
    $strings = $DomResult.strings 

    # 子フレームを含むすべてのドキュメントをループする。
    foreach ($doc in $DomResult.documents) {
        $nodes = $doc.nodes
        $layout = $doc.layout

        # ノード配列の要素数を特定する。
        $count = 0
        if ($nodes.parentIndex) { $count = $nodes.parentIndex.Count }
        elseif ($nodes.nodeName) { $count = $nodes.nodeName.Count }
        
        # ----------------------------------------------------------
        # 1. DOM構造の展開： CDPの圧縮配列を展開し、要素ごとのオブジェクトを生成する。
        for ($i = 0; $i -lt $count; $i++) {

            # 1要素分の器（テンプレート）を作成する。
            $node = [ordered]@{
                index          = $i
                parentIndex    = if ($nodes.parentIndex) { $nodes.parentIndex[$i] } else { $null }
                nodeType       = if ($nodes.nodeType) { $nodes.nodeType[$i] } else { $null }
                backendNodeId  = if ($nodes.backendNodeId) { $nodes.backendNodeId[$i] } else { $null }
                childIndexes   = if ($nodes.childNodeIndexes) { $nodes.childNodeIndexes[$i] } else { $null }
                attributes     = @{}
                computedStyles = @{}
                tagName        = ""
                id             = ""
                name           = ""
                placeholder    = ""
                value          = ""
                ariaLabel      = ""
                role           = ""
                innerText      = ""
            }

            # --- タグ名 (nodeName) の復元 ---
            # nodeName配列には直接の文字列ではなく、$strings 配列への「インデックス番号」が入っている。
            if ($nodes.nodeName -and $null -ne $nodes.nodeName[$i]) {
                # 該当インデックス番号をもとに、辞書($strings)から実際のタグ名文字列を引き当てる。
                $nameIndex = $nodes.nodeName[$i]
                if ($nameIndex -ge 0 -and $nameIndex -lt $strings.Count) {
                    $node.tagName = $strings[$nameIndex].ToLower()
                }
            }

            # --- 属性 (attributes) の復元 ---
            <# attributes配列は [名前のインデックス, 値のインデックス, 名前のインデックス, 値のインデックス...] 
               という連続した1次元配列として格納されているため、2つずつ(Step 2)取り出して復元する。 #>
            if ($nodes.attributes -and $nodes.attributes[$i]) {
                $attrArray = $nodes.attributes[$i]
                for ($a = 0; $a -lt $attrArray.Count; $a += 2) {
                    $nameIndex  = $attrArray[$a]
                    $valueIndex = $attrArray[$a + 1]

                    if ($null -ne $nameIndex -and $null -ne $valueIndex) {
                        # $strings 配列から実際の文字列を引き当てる。
                        $attrName  = $strings[$nameIndex]
                        $attrValue = $strings[$valueIndex]

                        if ($attrName) {
                            # 連想配列(Hash)へ属性名と値のペアを追加する。
                            $node.attributes[$attrName] = $attrValue

                            # RPAの要素特定の要となる主要属性は、アクセスしやすいようルート階層にもコピーする。
                            if ($attrName -eq "id") { $node.id = $attrValue }
                            if ($attrName -eq "name") { $node.name = $attrValue }
                            if ($attrName -eq "placeholder") { $node.placeholder = $attrValue }
                            if ($attrName -eq "value") { $node.value = $attrValue }
                            if ($attrName -eq "aria-label") { $node.ariaLabel = $attrValue }
                            if ($attrName -eq "role") { $node.role = $attrValue }
                        }
                    }
                }
            }

            # テキストノード(nodeType=3)の抽出と親要素への結合
            if ($node.nodeType -eq 3 -and $nodes.nodeValue -and $null -ne $nodes.nodeValue[$i]) {
                $valIndex = $nodes.nodeValue[$i]
                if ($valIndex -ge 0 -and $valIndex -lt $strings.Count) {
                    $textVal = $strings[$valIndex].Trim()
                    if (-not [string]::IsNullOrWhiteSpace($textVal)) {
                        # 親ノード(AタグやSPAN等)へテキストを連結する
                        if ($null -ne $node.parentIndex -and $node.parentIndex -ge 0 -and $node.parentIndex -lt $list.Count) {
                            $parentNode = $list[$node.parentIndex]
                            if (-not $parentNode.innerText) {
                                $parentNode.innerText = $textVal
                            } else {
                                $parentNode.innerText += " " + $textVal
                            }
                        }
                    }
                }
            }

            $list.Add($node)
        }

        # ----------------------------------------------------------
        # 2. レイアウト(スタイル)情報の紐づけ： 要素に対応するスタイル配列を辞書から引き当てる。

        # （レイアウトツリーが存在し、スタイル配列がある場合のみ実行する）
        if ($layout -and $layout.nodeIndex -and $layout.styles) {

            # layout.nodeIndex には、画面に描画されている要素のインデックス番号が入っている。
            for ($layoutIdx = 0; $layoutIdx -lt $layout.nodeIndex.Count; $layoutIdx++) {
                $nodeIdx = $layout.nodeIndex[$layoutIdx]

                # list内の該当ノードを取得する (フレーム毎にオフセットを考慮)
                # $list は累積されているため、現在のフレームの開始位置からのインデックスを計算する。
                if ($null -eq $nodeIdx) { continue } # ← 安全対策

                $absoluteIdx = ($list.Count - $count) + $nodeIdx

                if ($absoluteIdx -ge 0 -and $absoluteIdx -lt $list.Count) {
                    $targetNode = $list[$absoluteIdx]
                    $styleValues = $layout.styles[$layoutIdx]
                    
                    if ($null -eq $styleValues) { continue } # ← 安全対策
                    # リクエストした StyleNames の順序に合わせて文字列辞書から抽出する。
                    for ($styleIdx = 0; $styleIdx -lt $styleValues.Count -and $styleIdx -lt $StyleNames.Count; $styleIdx++) {
                        $strIdx = $styleValues[$styleIdx]
                        # PowerShell特有の $null -ge 0 が True になる問題を防ぐ厳密チェック
                        if ($null -ne $strIdx -and $strIdx -ge 0 -and $strIdx -lt $strings.Count) {
                            # 指定されたスタイル名をキーとして、辞書から引き当てた値を設定する。
                            $targetNode.computedStyles[$StyleNames[$styleIdx]] = $strings[$strIdx]
                        }
                    }
                }
            }
        }
    }
    return $list
}

# --- DOMSnapshotとBoxModel座標の統合取得 ---
# インタラクティブな操作対象要素に限定し、DOM構造とBoxModel(物理座標)を統合取得して通信負荷を軽減する。(内部関数)
<#
  画面上の全要素に対して座標取得(DOM.getBoxModel)を行うと、数千回のCDP通信が発生しフリーズの原因となる。
  まず一括でDOM構造を取得(DOMSnapshot)し、自動化の操作対象となり得るインタラクティブ要素(input, button等)に
  絞って座標を取得し、高速化と正確性を両立させる。
#>
function Get-DomSnapshotWithBoxModel {

    $func = $MyInvocation.MyCommand.Name

    # 1. 画面全体のDOMツリーと基本スタイルを一撃で取得する（この時点では座標x,yは含まれない）。
    $cdpResponse = Invoke-CdpCommand -Method "DOMSnapshot.captureSnapshot" -Params @{
        computedStyles = @("display","visibility","width","height")
    }
    $domResult = $cdpResponse.result
    if (-not $domResult) {
        Write-DebugLog -Message "[$func] エラー: DOMSnapshotの取得結果が空です" -Level Error
        return $null
    }

    # 2. 圧縮されたCDPの1次元配列を、扱いやすいオブジェクトリスト（連想配列）に展開する。
    $nodes = Convert-DomSnapshotFlat -DomResult $domResult

    # 3. DOM.getBoxModelのViewport座標を絶対座標へ補正するため、スクロール量を取得する。
    $scrollX = 0; $scrollY = 0
    try {
        # windowのスクロール量をJSで取得する。
        $scrJs = "return JSON.stringify({x: window.scrollX || 0, y: window.scrollY || 0});"
        $scrRes = Invoke-CdpScript -Js $scrJs
        $scrObj = $scrRes | ConvertFrom-Json
        if ($scrObj) { $scrollX = $scrObj.x; $scrollY = $scrObj.y }
    } catch {}

    # 4. インタラクティブ要素に絞って絶対座標を付与する。
    foreach ($node in $nodes) {
        if (-not $node.tagName) { continue }
        # パフォーマンス最適化：標準のインタラクティブ要素を許可
        $isInteractive = ($node.tagName -in @("input","button","a","select","textarea","iframe"))
        
        # SPA/エディタ対策：標準タグ以外でも、操作可能な役割(role)やラベル、編集可能(contenteditable)な要素は許可
        if (-not $isInteractive) {
            if ($node.role -match "^(button|link|checkbox|menuitem|tab|searchbox|textbox)$" -or
                -not [string]::IsNullOrEmpty($node.ariaLabel) -or
                $node.attributes["contenteditable"] -eq "true" -or
                $node.tagName -eq "body") {
                $isInteractive = $true
            }
        }
         
        # 該当しない純粋なレイアウトタグ(ただのdivやspan)はスキップして通信負荷を防ぐ
        if (-not $isInteractive) { continue }

        # ターゲット要素の固有ID（backendNodeId）がない場合はスキップする。
        if (-not $node.backendNodeId) { continue }

        try {
            # 絞り込んだ要素に対してのみ、CDP経由で物理座標（BoxModel）を要求する。
            $boxResponse = Invoke-CdpCommand -Method "DOM.getBoxModel" -Params @{
                backendNodeId = $node.backendNodeId
            }
            $model = $boxResponse.result.model
            # content ではなく、getBoundingClientRect() と同じ border 領域を参照する。
            if ($model -and $model.border) {
                # Viewport座標にスクロール量を足して、ページ絶対座標にする。
                $node.boundingBox = @{
                    x      = $model.border[0] + $scrollX
                    y      = $model.border[1] + $scrollY
                    width  = $model.width
                    height = $model.height
                }
            }

        } catch {
            # 無言の理由： 画面に描画されない要素に対するCDPエラーを無視してログスパムを防止する。
        }
    }

    $bbCount = ($nodes | Where-Object { $_.boundingBox }).Count
    if ($global:IsDebugMode) {
        Write-DebugLog -Message "[$func] 情報: DOMSnapshot 抽出ノード数: $($nodes.Count) / boundingBox 取得数: $bbCount" -Level Info
    }
    return $nodes
}

# --- 画面レイアウト情報エクスポート( JSON保存 ) ---
# Page.getLayoutMetrics を用い、DPRやViewportを含むレイアウト情報を保存する。
function Export-WebLayoutDump {
    param([string]$Prefix = "LayoutDump")

    $func = $MyInvocation.MyCommand.Name
    if ($null -eq $global:LogDir -or -not (Test-Path $global:LogDir)) {
        Write-DebugLog -Message "[$func] 警告: ログディレクトリ未発見" -Level Warning
        return
    }

    try {
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $filePath  = Join-Path $global:LogDir "${Prefix}_${timestamp}.json"

        # CDPコマンドを呼び出してレイアウト情報を取得する。
        $result = Invoke-CdpCommand -Method "Page.getLayoutMetrics" -Params @{}

        if (-not $result) {
            throw (New-EngineException -Func $func -Type "CDPエラー" `
                -Message "LayoutMetrics の取得結果が空です" -Details "Invoke-CdpCommand の戻り値が null / 空でした")
        }

        # JSON文字列へ変換する。
        $json = $result | ConvertTo-Json -Depth 16
        # JSONデータをファイルとして保存する。
        try {
            Set-Content -Path $filePath -Value $json -Encoding UTF8
        }
        catch {
            throw (New-EngineException -Func $func -Type "ファイルエラー" `
                -Message "LayoutDump の保存に失敗しました" -Details $_.Exception.Message)
        }
        return "[OK] Export-WebLayoutDump: $filePath"

    } catch {
        # すでに New-EngineException で生成されたエラーならそのままスローする。
        if ($_.Exception.Message -match "^\[.*?\] \[.*?\]:") { throw $_ }
        
        throw (New-EngineException -Func $func -Type "内部エラー" `
            -Message "DOM Snapshot の取得処理中に例外が発生しました" -Details $_.Exception.Message)
    }
}

# ==============================================================================
# --- セレクタ安定性検証 (テスト) ---
# ==============================================================================
<#
  取得したセレクタが、動的ID(ext-gen-* 等)や仮想DOMの再描画によって短時間で消滅・変化しないか(一過性の幻ではないか)を複数回検証し、
  RPA本番実行時の「要素が見つからない」エラー(Flakyテスト)を未然に防ぐ。
#>
function Test-SelectorStability {
    param(
        [string]$Selector,
        [int]$Repeat = 3
    )

    $func = $MyInvocation.MyCommand.Name
    Write-DebugLog -Message "=== Selector 安定性テスト ===／ [TARGET] = $Selector" -Level Info

    if (-not $global:WebViewCtrl) {
        Write-DebugLog -Message "[$func] [NG] WebViewCtrl が未初期化です" -Level Error
        return
    }

    # Base64エンコードを利用してサニタイズする。
    $selB64 = ConvertTo-JsSafeBase64 -Text $Selector

    $js = @"
    return (function(){
        try {
            // Base64化されたセレクタをデコードして復元する。
            const sel = decodeURIComponent(escape(atob('$selB64')));
            // 共通関数(utilFindInFrames)を利用し、iframe内も含めて要素を透過探索する。
            const el = utilFindInFrames(window, function(win) {
                try {
                    // プレフィックスがXPATH:の場合はXPath評価、それ以外はquerySelectorを実行する。
                    return sel.indexOf('XPATH:') === 0
                        ? win.document.evaluate(sel.substring(6), win.document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue
                        : win.document.querySelector(sel);
                } catch(e) { return null; }
            });

            // 要素が発見できなかった場合はexistsフラグをfalseにして返却する。
            if (!el) {
                return JSON.stringify({ exists: false });
            }

            // 発見できた場合は、要素の座標(BoundingBox)と属性を取得しJSONで返却する。
            const rect = el.getBoundingClientRect();
            return JSON.stringify({
                exists: true,
                tag: el.tagName,
                id: el.id || "",
                name: el.name || "",
                rect: {
                    left: Math.round(rect.left),
                    top: Math.round(rect.top),
                    width: Math.round(rect.width),
                    height: Math.round(rect.height)
                }
            });
        } catch(e) {
            return JSON.stringify({ exists: false, error: e.message });
        }
    })();
"@

    # 指定された回数(Repeat)だけ、一定間隔で安定して要素が取得できるかテストを繰り返す。
    for ($i = 1; $i -le $Repeat; $i++) {
        # JSを実行し、JSONをPowerShellのオブジェクトへデコードする。
        $res = Invoke-WebScript -Js $js
        $info = if ($res -is [string]) { 
            try { $res | ConvertFrom-Json -ErrorAction Stop } 
            catch { throw (New-EngineException -Func $func -Type "内部エラー" -Message "セレクタテスト結果のJSONデコードに失敗しました" -Details $_.Exception.Message) } 
        } else { $res }

        if (-not $info) {
            Write-DebugLog -Message "[$func] [RUN $i/$Repeat]: [NG] JS結果の解析に失敗" -Level Error
            continue
        }

        if (-not $info.exists) {
            Write-DebugLog -Message "[$func] [RUN $i/$Repeat]: [NG] 要素が見つかりません (Selector=$Selector)" -Level Error
            if ($info.error) {
                Write-DebugLog -Message "      Error: $($info.error)" -Level Error
            }
            continue
        }
        Write-DebugLog -Message ("[RUN $i/$Repeat]: [OK] tag=$($info.tag), id=$($info.id), name=$($info.name) " +
                         ",[left=$($info.rect.left), top=$($info.rect.top), width=$($info.rect.width), height=$($info.rect.height)]") -Level Success
    }
}
