# ------------------------------------------------------------------------------#●#✘ 
<#
  [WebView2 ネイティブ通信モジュール]： WebView2標準の ExecuteScriptAsync API を利用したJavaScript実行基盤を定義する。
  CDP通信が制限された環境や、CDP非対応時のメイン制御・フォールバックとして機能する。
#>
# ------------------------------------------------------------------------------

# --- WebView2標準APIを利用したJavaScript非同期実行と結果取得 ---
# WebView2標準の ExecuteScriptAsync APIを利用し、JSの非同期実行と結果取得を行う。
function Invoke-WebView2NativeScript {
    param (
        [Parameter(Mandatory = $true)][string]$Js,
        [int]$Retries = 3,
        [int]$InitialDelayMs = 200
    )

    $func = $MyInvocation.MyCommand.Name
    $jsCode = $Js

    # JSの安全なラップおよびJSON形式での返却： 異言語間での型とエラー状態を維持する。
    # 渡されたJSコードを即時実行関数(IIFE)で包み、例外発生時のスタックトレースをPowerShell側へ引き渡す。
    $wrappedJs = @"
(function() {
    try {
        // 渡されたコード(jsCode)を即時実行関数内で実行し、戻り値を変数resultへ格納する。
        const result = (function() { $jsCode })();

        // undefinedをnullに置換してJSON消失を防止する。
        return JSON.stringify({
            status: "success",
            data: result !== undefined ? result : null
        });
    } catch (e) {
        // 例外発生時は、エラー情報(メッセージ、スタックトレース、エラー名)を抽出してJSON文字列化する。
        return JSON.stringify({
            status: "error",
            message: e && e.message ? e.message : String(e),
            stack: e && e.stack ? e.stack : "",
            name: e && e.name ? e.name : ""
        });
    }
})();
"@

    $attempt = 0
    $delayMs = $InitialDelayMs

    # 成功するかリトライ上限に達するまで無限ループを実行する。
    while ($true) {
        $attempt++

        try {
            # アクティブタブのWebViewコントロール取得： 実行対象となる現在のタブを取得する。
            $webview = Get-ActiveWebView
            if (-not $webview) {
                throw (New-EngineException -Func $func -Type "初期化エラー" -Message "アクティブなWebViewの取得に失敗しました（タブ未初期化）")
            }

            # JSの非同期実行： ExecuteScriptAsyncを利用してJavaScriptをブラウザ側へ送信する。
            $task = $webview.CoreWebView2.ExecuteScriptAsync($wrappedJs)

            # 非同期処理の完了待機： UIスレッドのフリーズを防ぐため、完了するまでDoEventsを回して待機する。
            while (-not $task.IsCompleted) {
                [System.Windows.Forms.Application]::DoEvents()
                Start-Sleep -Milliseconds 10
            }

            if ($task.IsFaulted) {
                throw (New-EngineException -Func $func -Type "ネイティブエラー" -Message "WebView2ネイティブAPI(ExecuteScriptAsync)の実行に失敗しました" -Details $task.Exception.InnerException.Message)
            }

            # 正常終了した非同期タスクから、実行結果(文字列)を取り出す。
            $rawResult = $task.Result

            # undefinedの評価と変換： undefinedをPowerShellのnullとして評価・変換する。
            <# 
              JS側で関数が値を返さない(undefined)場合、文字列として返却されてしまうため、
              JSON変換でのキー消失を防ぐべく、PowerShell側の $null として安全に評価・変換する。
            #>
            if ($rawResult -eq '"undefined"' -or $rawResult -eq 'undefined') {
                return $null
            }

            # 空またはnullの場合のリトライ処理： 指数バックオフを用いてページ遷移直後などの無応答状態を救済する。
            if ([string]::IsNullOrWhiteSpace($rawResult) -or $rawResult -eq "null") {
                # 試行回数が上限に達していないか判定する。
                if ($attempt -lt $Retries) {
                    Start-Sleep -Milliseconds $delayMs
                    # 次回の待機時間を2倍に設定する。(上限2000ミリ秒の指数バックオフ)
                    $delayMs = [Math]::Min(2000, $delayMs * 2)
                    continue
                } else {
                    throw (New-EngineException -Func $func -Type "内部エラー" -Message "WebView2ネイティブAPIの戻り値が空またはnullです")
                }
            }

            # ExecuteScriptAsync 特有の二重エスケープの解除： 戻り値に付与される不要なダブルクォートとエスケープを剥がす。
            <# 
              戻り値が文字列の場合、先頭と末尾に不要なダブルクォートが付与され、
              内部がエスケープされるWebView2の仕様があるため、それを剥がして復元する。
            #>
#✘         if ($rawResult -match '^\s*"(.*)"\s*$') {
            # 取得テキストに改行が含まれる場合を考慮し、正規表現に単一行モード（(?s)）を付与する。
            if ($rawResult -match '(?s)^\s*"(.*)"\s*$') {
                # マッチした内部の文字列($Matches[1])を抽出する。
                $inner = $Matches[1]
                try {
                    # .NETのRegex.Unescapeを利用し、\n や \" などのエスケープ文字を純粋な文字列へデコードする。
                    $rawResult = [System.Text.RegularExpressions.Regex]::Unescape($inner)
                } catch {
                    $rawResult = $inner
                }
            }

            # JSONのデコード処理： 文字列化されたJSの実行結果をPowerShellのオブジェクトへ変換する。
            try {
                $response = $rawResult | ConvertFrom-Json -ErrorAction Stop
            } catch {
                if ($attempt -lt $Retries) {
                    Start-Sleep -Milliseconds $delayMs
                    $delayMs = [Math]::Min(2000, $delayMs * 2)
                    continue
                } else {
                    throw (New-EngineException -Func $func -Type "内部エラー" -Message "WebView2からの応答(JSON)のデコードに失敗しました")
                }
            }

            # JS側で発生した例外のエラー処理： スタックトレース等を抽出し、PowerShell側の例外として再送出する。
            if ($response.status -ne "success") {
                $stack = $response.stack -replace "`r`n", " | "
                throw (New-EngineException -Func $func -Type "JSエラー" -Message "$($response.name): $($response.message)" -Details $stack)
            }

            return $response.data

        } catch {
            # 実行時エラーに対する全体リトライ処理： 予期せぬ実行エラー発生時も上限回数までリトライを試行する。
            if ($attempt -lt$Retries) {
                Start-Sleep -Milliseconds $delayMs
                $delayMs = [Math]::Min(2000, $delayMs * 2)
                continue
            } else {
                $errMsg = (New-EngineException -Func $func -Type "ネイティブエラー" -Message "スクリプトの実行処理中に予期せぬエラーが発生しました" -Details $_.Exception.Message)
                throw $errMsg
            }
        }
    }
}
