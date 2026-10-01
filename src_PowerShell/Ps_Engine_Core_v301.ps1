# ------------------------------------------------------------------------------#●#✘ 
# 自立型 WebView2 RPAエンジン（Core）
# コマンド受信および各モジュールへのルーティング
# advice by AI (2026/05 - /09/..)
# ------------------------------------------------------------------------------
<#
    ============================================================================
• 命名規則およびコーディングルール
• 実行環境: PowerShell 5.1
    ============================================================================
【1. 変数およびパラメータの基本命名規則】
• ローカル変数: camelCase（キャメルケース）
   ※ ローカル変数の解釈拡大:
      関数内の変数だけでなく、「モジュール展開時（ドットソース）の初期化処理でのみ使用し、
      以降の状態として参照しない作業用変数」もローカル変数と同義とする。
   ※ パラメータ名（PascalCase）をローカルで加工・再利用する際は、必ず camelCase の別名とする。
      [NG]: $XPathEscaped = Normalize-XPath $XPath
      [OK]: $xpathEscaped = Normalize-XPath $XPath

• パラメータおよびグローバル変数: PascalCase（パスカルケース）
• 定数・システム設定値（例外ルール）:
   $global:CONFIG のように、定数・設定情報として機能する変数は大文字を許容する。

【2. 時間・タイムアウト系の命名規則（厳格化）】
• 時間や間隔を指定する変数は、必ず「単位（Sec/Ms）」を接尾辞として明記する。
   ※ Timeout（単位不明）等の曖昧な命名は使用禁止。
   - 秒単位の例: TimeoutSec, WaitTimeSec
   - ミリ秒単位の例: TimeoutMs, PollIntervalMs, InitialDelayMs, WaitMs
   ※ これらも「基本命名規則」に従い、パラメータなら $TimeoutSec、ローカルなら $timeoutSec となる。

【3. 異言語間（VBA / PowerShell / JavaScript）連携の命名規則】
• [VBA → PowerShell] 真偽値（Boolean）のコマンドライン引数渡し:
   プロセス間通信において、真偽値を文字列型（"true"/"false"）で受け取る起動引数（param）には、
   内部で保持する Boolean 変数と区別するため、意図的に Flg という接尾辞を付与する。
   - 例: [string]$IsDebugModeFlg

• [PowerShell → JavaScript] 埋め込み用変数の統一:
   PowerShellからJSのヒアドキュメント（@""@）に展開・埋め込むための加工済み変数は、
   元のパラメータが PascalCase であっても、必ず camelCase に変換して埋め込む。
   - 例: $XPath (パラメータ) → $xpathEscaped (JS埋め込み用ローカル変数)
   - 例: $FrameSelector → $frameSelectorEscaped

• [JavaScript 内部コード] 変数名の独立と統一:
   JSコード内で宣言する変数は、JavaScriptの標準規約に従い camelCase を徹底する。
   PowerShell側の PascalCase をJSコンテキスト内に持ち込まない。
   - [NG]: const XPath = '$xpathEscaped';
   - [OK]: const xpath = '$xpathEscaped';
#>
<#
    ============================================================================
■ 外部関数 (External / Public API)
   VBA側の `rpaEngine.RunAction` から直接 `Command` として呼び出されることを想定した公開インターフェース
   (例: Invoke-WebClick, Set-WebTextInput, Enable-SilentDownload 等)
    ============================================================================
■ 内部関数 (Internal Functions)    
   [ログ・例外ハンドリング]
   - Write-DebugLog        : 実行ログ、デバッグ情報のコンソール出力およびファイル書き込み
   - New-EngineException   : 例外情報の標準フォーマット化（VBAのParseRpaErrorと対をなす）
    
   [待機・同期ユーティリティ]
   - Wait-Condition        : 任意の条件スクリプトブロックが $true になるまでの汎用ポーリング待機
    
   [文字列・セレクタ解析]
   - Normalize-XPath       : HTML上の改行や空白による不一致を防ぐためのXPath自動補正
    
   [コア通信・セッション制御]
   - Get-ActiveWebView     : 現在操作対象のアクティブなWebView2インスタンスを取得
   - Connect-CdpSession    : CDP (Chrome DevTools Protocol) のWebSocket接続を確立・維持
   - Invoke-WebScript      : JS実行の最上位窓口 (Native/CDPのルーティングと型統一ラッパーを適用)
       ├─ Invoke-CdpScript           : (内部) CDP経由でのJS評価
       └─ Invoke-WebView2NativeScript: (内部) Native(ExecuteScriptAsync)でのJS評価
   
   ◆ Lib-WebJS に移設して、Lib-WebView2_Init でグローバル注入する。
   - deepQuerySelector     : (JS関数) Shadow DOM 貫通検索
   - utilFindInFrames      : (JS関数) iframe透過探索
#>
<#
    ============================================================================
【JavaScript インジェクションに関する安全基準】
1. 変数宣言のモダン化 (グローバルスコープ汚染の防止)
   - WebView2 (Chromium) へ注入・実行するJSコード内では、従来の `var` による変数宣言を原則禁止とする。
   - 意図せぬ状態の共有やバグを防ぐため、再代入が不要な変数は `const`、必要な変数は `let` に統一する。
2. パラメータの完全サニタイズ (Base64 カプセル化による構文破壊防止)
   - PowerShellからJSへ変数を埋め込む際、シングルクォート等の単純な文字列置換（.Replace("'", "\'")）は
     XSS的な構文破壊の要因となる。
   - 確実な安全性を担保するため、文字列は必ず `ConvertTo-JsSafeBase64` でエンコードし、
     JS側にて `decodeURIComponent(escape(atob('$varB64')))` で復元するカプセル化方式を標準とする。
#>
<#
    ============================================================================
【インラインコメント、(マーカー)】 **このRPAルール
• 処理の意図や背景が明確化 (「なぜ必要なのか」「何をしているのか」)
• 関数 及び 見出しは「体言止め」、補足説明は「〜する。」などの「常体（だ・である調）」を基本とする。
• 関数（function）は --- 見出し --- とする。大ブロックは 括弧 [見出し] とする。
• 大ブロックの解釈】 括弧 [ ] を用いた大ブロック指定は、ルートレベル（モジュール読み込みやメインループ等、大きな処理の塊）にのみ適用する。
 【関数内、ループ内、イベントハンドラ内の個別処理からは [ ] を外し、通常のインラインコメントに降格させることで視覚的な階層を明確にする。
• インデックス（見出し）と補足説明が混在する場合は、コロン（：）と半角スペースで区切る。 （基本フォーマット: 見出し（体言止め）： 補足説明（〜する。））
• 各コードの直上に付けるコメントは1行を基本とし、見出しを伴わない単独の説明文の場合は「〜する。」などの「常体」とする。
• ただし、処理の意図などコメントが長くなる場合のみ、改段、及び《# ... #》（ブロックコメント）で説明文を記述する。

・#● デバッグ用（開発時には、有った方が便利な箇所）
・#✘ デバッグ用（開発時に不要コード、修正前としたコード）
#>


# ------------------------------------------------------------------------------
# [起動引数の受け取り]： 外部(VBA)から渡される初期設定値を受け取る。
param (
    [string]$AppSessionId = "AUTO_$(Get-Date -Format 'yyyyMMddHHmm')",
    [int]$CdpPort = 0,
    [string]$IsDebugModeFlg = "true"
)

# [起動引数のグローバル保持]： 各モジュールから状態を共有・参照できるよう、起動引数をグローバル変数へ格納する。
$global:AppSessionId = $AppSessionId
$global:CdpPort = $CdpPort
try {
    $global:IsDebugMode = [bool]::Parse($IsDebugModeFlg)
} catch {
    $global:IsDebugMode = $true
}

# ------------------------------------------------------------------------------
# [UIアセンブリのロードおよびDPI基準統一]： DPIの基準を強制的に統一する。
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Drawing

<#
    ============================================================================
  【DPIスケーリングの強制発火ハック】
  WPF(System.Windows.Automation)は、初期化されるまでOSの現在のDPIスケーリング(125%や150%など)を正しく認識しません。
  このため、自動化処理の直前にダミーの座標(0,0)オブジェクトを生成してUIA要素を強制取得させることで、
  WPFの内部エンジンを意図的に初期化し、以降の物理マウスクリック等における座標計算のズレを完全に防ぎます。
    ============================================================================
#>
try {
    # 画面サイズを取得する前にWPFのシステムを初期化し、DPIを真の物理サイズに固定する。
    $dummyPoint = New-Object System.Windows.Point(0,0)
    # オブジェクトからUIAutomation要素を取得し、WPFのDPIスケーリングを強制発火させる。
    [System.Windows.Automation.AutomationElement]::FromPoint($dummyPoint) | Out-Null
} catch {}

# [プライマリモニター作業領域の取得]： RPAの確実な描画と座標計算のため、タスクバー等を除外した真の有効領域サイズを取得する。
# OSのメインディスプレイ幅を取得して変数へ格納する。
$screenWidth  = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Width
# OSのメインディスプレイ高さを取得して変数へ格納する。
$screenHeight = [System.Windows.Forms.Screen]::PrimaryScreen.WorkingArea.Height

# ------------------------------------------------------------------------------
# [外部ウィンドウ制御用Win32 APIの宣言]： 起動時のフォーカス確保や、拡張モジュールでのマウス・ウィンドウ操作に利用する。
$win32Signature = @"
using System;
using System.Runtime.InteropServices;

namespace Win32Api {
    [StructLayout(LayoutKind.Sequential)]
    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    public class Win32Utils {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);
        
        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        public static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);
    }
}
"@

# [Win32 APIクラスの二重ロード防止と動的コンパイル]： Win32Api.Win32Utils 型が未定義の場合のみC#コードをコンパイルしてロードする。
if (-not ([System.Management.Automation.PSTypeName]"Win32Api.Win32Utils").Type) {
    # 未定義の場合のみ、C#のコード文字列をコンパイルしてアセンブリとしてロードする。
    Add-Type -TypeDefinition $win32Signature
}

# ------------------------------------------------------------------------------
# [共通設定値の初期化]： システム全体の動作を定義する設定辞書を生成する。
$global:CONFIG = @{
    MaxLogGenerations = 5
    DefaultTimeoutSec = 10
    BrowserWidth      = $screenWidth
    BrowserHeight     = $screenHeight
    EnableHighlight   = $true
}

# [通信モードに応じたポーリング間隔の設定]： 通信モードによりUI待機用のポーリング間隔を切り替える。
if ($global:CdpPort -gt 0) {
    # 高速ポーリング（CDPモード）を設定する。
    $global:CONFIG.PollIntervalMs = 50
} else {
    # UIスレッド負荷軽減（ネイティブモード）として、CPU負荷を抑える。
    $global:CONFIG.PollIntervalMs = 200
}

# [システム用グローバル変数の初期化]： ウィンドウやタブ管理用の変数を $null で定義する。
# ブラウザを配置するWinFormsウィンドウの格納先を初期化する。
$global:BrowserForm = $null
# ブラウザ本体となるWebView2コントロールの格納先を初期化する。
$global:WebViewCtrl = $null

# [タブ管理構造体の初期化]： 仮想タブ管理用のハッシュテーブルを定義する。
# 仮想タブやポップアップを管理するためのハッシュテーブルを初期化する。
$global:Tabs = @{}
# 操作対象となっているタブのIDを保持する変数を初期化する。
$global:ActiveTabId = $null

# ------------------------------------------------------------------------------
# [実行フォルダパスの取得]： スクリプトが配置されているディレクトリパスを取得する。
$global:ScriptDirectory = $PSScriptRoot
if ([string]::IsNullOrEmpty($global:ScriptDirectory)) {
    $global:ScriptDirectory = Split-Path -Parent $MyInvocation.MyCommand.Definition
}

# [ログフォルダの世代管理]： 指定された世代数を超える古いログフォルダを削除する。
$global:ParentLogDirectory = Join-Path $global:ScriptDirectory "Logs"
if (Test-Path $global:ParentLogDirectory) {
    $oldLogFolders = Get-ChildItem -Path $global:ParentLogDirectory -Directory | Sort-Object CreationTime
    if ($oldLogFolders.Count -ge $global:CONFIG.MaxLogGenerations) {
        $deleteCount = $oldLogFolders.Count - $global:CONFIG.MaxLogGenerations + 1
        $oldLogFolders | Select-Object -First $deleteCount | ForEach-Object {
            try { Remove-Item $_.FullName -Recurse -Force -ErrorAction Stop } catch {}
        }
    }
}

# [本セッション用ログフォルダの作成]： セッション用のログフォルダを作成する。
$global:LogDir = Join-Path $global:ParentLogDirectory $AppSessionId
if (-not (Test-Path $global:LogDir)) { 
    New-Item -ItemType Directory -Path $global:LogDir -Force | Out-Null
}

# [コンソール文字コードのShift-JIS強制]： VBAとの通信やログ出力時の文字化けを防ぐ。
try {
    $sjis = [System.Text.Encoding]::GetEncoding("Shift-JIS")
    [Console]::OutputEncoding = $sjis
    [Console]::InputEncoding  = $sjis
} catch {}

# --- ログ出力 ---
# 実行ログ、デバッグ情報のコンソール出力およびファイル書き込みを行う。
function Write-DebugLog {
    param (
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet('Info', 'Success', 'Warning', 'Error', 'Fatal')][string]$Level = 'Info'
    )

    # タイムスタンプとログレベルを付与する。
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $formattedMessage = "<$timestamp> <$Level> $Message"

    # コンソールへ出力する。
    try { [Console]::WriteLine("[LOG] $formattedMessage") } catch {}

    # ログファイルへ出力する。
    if ($null -ne $global:LogDir -and (Test-Path $global:LogDir)) {
        $logFilePath = Join-Path $global:LogDir "Engine_SystemLog.txt"
        try {
            $formattedMessage | Out-File -FilePath $logFilePath -Append -Encoding UTF8 -ErrorAction Stop
        } catch {
            # フェイルセーフ対応： OS標準Tempディレクトリへ退避し、プロセスのクラッシュを防ぐ。
            try {
                $tempLogPath = Join-Path ([System.IO.Path]::GetTempPath()) "RPA_Engine_FatalLog.txt"
                $fallbackMsg = "[FALLBACK] $formattedMessage"
                $fallbackMsg | Out-File -FilePath $tempLogPath -Append -Encoding UTF8
                
                # コンソール(VBA側)へ通知し、ログの退避先パスを把握可能にする。
                [Console]::WriteLine("[ERROR] ログファイルへの書き込みに失敗し、Tempフォルダへ退避しました: $tempLogPath")
            } catch {}
        }
    }
}

# --- 共通例外エラーフォーマット生成 ---
# VBAプロセスへ返す例外文字列を標準フォーマットで生成する。
function New-EngineException {
    param(
        [Parameter(Mandatory=$true)][string]$Func,
        [Parameter(Mandatory=$true)]
        [ValidateSet("Timeout","未発見","JSエラー","CDPエラー","ネイティブエラー","初期化エラー","ファイルエラー","UIAエラー","引数エラー","内部エラー")]
        [string]$Type,
        [Parameter(Mandatory=$true)][string]$Message,
        [string]$Details
    )

    # エラーログの多行跨ぎ防止および1行フォーマット正規化： 改行をスペースに置換し、ローカル変数へ格納する。
    $messageCleaned = $Message -replace "`r`n|`n|`r", " "
    
    if ([string]::IsNullOrWhiteSpace($Details)) {
        return "[$Func] [$Type]: $messageCleaned"
    } else {
        $detailsCleaned = $Details -replace "`r`n|`n|`r", " | "
        return "[$Func] [$Type]: $messageCleaned ($detailsCleaned)"
    }
}

# ------------------------------------------------------------------------------
# [DPR（ディスプレイ拡大率）の取得]： 座標計算に使用するOSレベルのスケーリング倍率を算出する。
$graphics = [System.Drawing.Graphics]::FromHwnd([IntPtr]::Zero)
$dpr = [math]::Round($graphics.DpiX / 96.0, 2)
$graphics.Dispose()

# 起動情報ログを出力する。
Write-DebugLog -Message "[System] 情報: Browser/ Width-Height ($screenWidth) - ($screenHeight) | DPR: $dpr" -Level Info
Write-DebugLog -Message "[System] 情報: 開発モードスイッチ ($global:IsDebugMode)" -Level Info
Write-DebugLog -Message "[System] 情報: 通信モードスイッチ ($global:CdpPort)" -Level Info

# ------------------------------------------------------------------------------
# [常時ロード対象モジュールの定義]： エンジン稼働に必要な基幹・拡張モジュール群を一括で読み込む。
$alwaysLoadLibs = @(
    # 基幹モジュールを定義する。
    "Lib-WebJS_v301.ps1",
    "Lib-WebView2_Init_v301.ps1",
    "Lib-WebView2_Native_v301.ps1",

    # 拡張モジュールを定義する。
    "Lib-WebAction_v301.ps1",
    "Lib-WebXPath_v301.ps1",
    "Lib-WebDebug_v301.ps1",
    "Lib-DesktopUIA_v301.ps1",
    "Lib-WebSafeAction_v301.ps1",
    "Lib-DevTools_v301.ps1"
)

foreach ($libName in $alwaysLoadLibs) {
    $libPath = Join-Path $global:ScriptDirectory $libName
    if (Test-Path $libPath) {
        try {
            . $libPath
            Write-DebugLog -Message "[System] 成功: モジュールをロード ($libName)" -Level Success
        } catch {
            Write-DebugLog -Message "[System] 致命的エラー: $libName のロードに失敗しました ($($_.Exception.Message))" -Level Fatal
            Exit
        }
    } else {
        Write-DebugLog -Message "[System] 致命的エラー: 必須モジュール $libName が見つかりません。起動を中止します。" -Level Fatal
        Exit
    }
}

# [CDPモード専用モジュールのロード]： CDP通信が有効な場合のみモジュール読み込みとポート接続確認を行う。
if ($global:CdpPort -gt 0) {
    $cdpLibName = "Lib-WebCDP_v301.ps1"
    $cdpLibPath = Join-Path $global:ScriptDirectory $cdpLibName

    if (Test-Path $cdpLibPath) {
        . $cdpLibPath
        Write-DebugLog -Message "[System] 成功: モジュールをロード ($cdpLibName)" -Level Success

        # CDPポート状態の確認： バインド遅延を考慮したリトライを実施する。
        $portReady = $false
        for ($i = 0; $i -lt 6; $i++) {
            Start-Sleep -Milliseconds 500
            try {
                $tcpConn = Get-NetTCPConnection -LocalPort $global:CdpPort -ErrorAction Stop | Select-Object -First 1
                Write-DebugLog -Message "[System] 情報: CDPポート ($global:CdpPort) の状態 - $($tcpConn.State) ($($tcpConn.LocalAddress))" -Level Info
                $portReady = $true
                break
            } catch {}
        }
        
        if (-not $portReady) {
            Write-DebugLog -Message "[System] 致命的エラー: CDPポート ($global:CdpPort) のリスン確認に失敗しました。WebView2プロセスの起動遅延または競合の可能性があります。" -Level Fatal
            Exit
        }
    }
}

# ------------------------------------------------------------------------------
# [タブ管理]： タブの取得、切り替え、一覧出力に関する処理群を定義する。

# --- アクティブWebView2インスタンス取得 ---
# 現在操作対象のアクティブなWebView2インスタンスを取得する。
function Get-ActiveWebView {
    # Tabs登録済みWebViewを返却する。
    if ($null -ne $global:ActiveTabId -and $global:Tabs.ContainsKey($global:ActiveTabId)) {
        return $global:Tabs[$global:ActiveTabId].WebView
    }
    # Tabs未登録時の後方互換として処理する。
    return $global:WebViewCtrl
}

# --- アクティブタブ設定 ---
# タブIDを直接指定してアクティブタブを切り替える（前面化）。
function Set-ActiveTab {
    param ([string]$TabId)

    $func = $MyInvocation.MyCommand.Name

    if ($global:Tabs.ContainsKey($TabId)) {
        $global:ActiveTabId = $TabId
        $global:Tabs[$TabId].WebView.BringToFront()
        return "ActiveTab = $TabId"
    }
    throw (New-EngineException -Func $func -Type "引数エラー" -Message "指定されたタブIDが見つかりません" -Details $TabId)
}

# --- タブ一覧取得 ---
# 起動中の全タブ情報（ID、URL、タイトル）をJSONで取得する。
function List-Tabs {
    $result = @()

    foreach ($tabId in $global:Tabs.Keys) {
        $tab = $global:Tabs[$tabId]
        $webview = $tab.WebView

        # Tabs連想配列URLの優先取得： CDP同期用として優先的に取得する。
        $url = ""
        if ($tab.ContainsKey("Url") -and $tab.Url) {
            $url = $tab.Url
        } else {
            try { $url = $webview.Source.ToString() } catch {}
        }

        # WebView2からタイトルを取得する。
        $title = ""
        try { $title = $webview.CoreWebView2.DocumentTitle } catch {}

        $result += [PSCustomObject]@{
            TabId    = $tabId
            Url      = $url
            Title    = $title
            IsActive = ($tabId -eq $global:ActiveTabId)
        }
    }
    Write-DebugLog -Message "[List-Tabs] タブ数: $($result.Count)" -Level Info
    # JSON変換時のパイプ回避： パイプ (|) を使わず、-InputObject に直接配列を渡す。
    $resultJson = ConvertTo-Json -InputObject @($result) -Depth 2 -Compress
    Write-DebugLog -Message "$resultJson" -Level Info
    return $resultJson
}

# --- タブ切替 ---
# タブIDによる切り替えを行う（CDPの再接続処理も包含する）。
function Switch-Tab {
    param ([Parameter(Mandatory=$true)][string]$TabId)

    $func = $MyInvocation.MyCommand.Name

    if (-not $global:Tabs.ContainsKey($TabId)) {
        throw (New-EngineException -Func $func -Type "引数エラー" -Message "指定されたタブIDが見つかりません" -Details $TabId)
    }

    # アクティブタブを更新する。
    $global:ActiveTabId = $TabId

    # WebView2を前面に表示する。
    $webview = $global:Tabs[$TabId].WebView
    try {
        $webview.BringToFront()
        $global:WebViewCtrl = $webview
    } catch {}

    if ($global:CdpPort -gt 0) {
        try {
            # Connect-CdpSession -Port $global:CdpPort
        } catch {
            Write-DebugLog -Message "[Switch-Tab] CDP再接続失敗: $($_.Exception.Message)" -Level Warning
        }
    }
    return "Switched to Tab: $TabId"
}

# --- タイトル部分一致によるタブ切替 ---
# タイトルの部分一致検索によるタブ切り替えを行う。
function Switch-TabByTitle {
    param ([Parameter(Mandatory=$true)][string]$TitleSubstring)
    
    $func = $MyInvocation.MyCommand.Name
    
    foreach ($tabId in $global:Tabs.Keys) {
        $title = ""
        try { 
            # WebView2コントロールからタイトルを取得する。
            $title = $global:Tabs[$tabId].WebView.CoreWebView2.DocumentTitle
        } catch {}
        
        if ($title -like "*$TitleSubstring*") {
            Write-DebugLog -Message "[$func] 情報: タブ切り替え ($tabId - $title)" -Level Info
            return Switch-Tab -TabId $tabId
        }
    }
    # 既存の例外フォーマットへ統一する。
    throw (New-EngineException -Func $func -Type "未発見" -Message "指定されたタイトルを含むタブが見つかりません" -Details $TitleSubstring)
}

# ------------------------------------------------------------------------------
# [汎用ロジックおよび JS 実行ルーター]： 待機処理やJSのハイブリッドルーティング機能を提供する。

# --- タイムアウト付き汎用待機 ---
# UIフリーズを防止しつつ、指定条件がTrueになるまで待機する（汎用）。
function Wait-Condition {
    param (
        [Parameter(Mandatory=$true)][scriptblock]$ConditionBlock,
        [int]$TimeoutSec = $global:CONFIG.DefaultTimeoutSec,
        [int]$PollIntervalMs = $global:CONFIG.PollIntervalMs,
        [string]$TimeoutMessage = "待機処理がタイムアウトしました"
    )

    $func = $MyInvocation.MyCommand.Name

    # UIフリーズを防止して条件成立を待機する。
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        if (& $ConditionBlock) { return $true }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds $PollIntervalMs
    }
    throw (New-EngineException -Func $func -Type "Timeout" -Message $TimeoutMessage)
}

# --- JavaScript実行ルーティング ---
# JS実行のルーティングを行う（CDPが有効ならCDP、失敗時はNativeへフォールバックする）。
function Invoke-WebScript {
    param (
        [Parameter(Mandatory=$true)][string]$Js,
        [int]$Retries = 3
    )

    # アクティブWebViewを取得する。
    $webview = Get-ActiveWebView

    # CDP通信により最優先で実行する。
    if ($null -ne $global:CdpWebSocket -and $global:CdpWebSocket.State -eq 'Open') {
        try {
            return Invoke-CdpScript -Js $Js
        } catch {
            Write-DebugLog -Message "[Engine] 警告: CDP実行失敗、ネイティブ実行へフォールバック ($($_.Exception.Message))" -Level Warning
        }
    }
    # ネイティブ実行へフォールバックする。
    return Invoke-WebView2NativeScript -Js $Js -Retries $Retries
}

# --- JSインジェクション安全化用Base64エンコード ---
# PowerShellからJSへ変数を安全に埋め込むため、文字列をBase64にエンコードする。
function ConvertTo-JsSafeBase64 {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return "" }
    return [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($Text))
}

# --- エンジン設定の動的変更 ---
# 実行時のエンジン設定（要素ハイライトのON/OFF等）を動的に変更する。
function Set-EngineConfig {
    # ハイライト機能設定： VBA側からのグローバル設定オブジェクトを更新する。
    param ([string]$EnableHighlight)

    $func = $MyInvocation.MyCommand.Name

    # EnableHighlightをBoolean変換して更新する。
    if (-not [string]::IsNullOrEmpty($EnableHighlight)) {
        $global:CONFIG.EnableHighlight = [System.Convert]::ToBoolean($EnableHighlight)
#●        Write-DebugLog -Message "[$func] 設定変更: EnableHighlight = $($global:CONFIG.EnableHighlight)" -Level Info
    }
    return "Config Updated"
}

# ==============================================================================
# [通信ループおよびコマンド待機]： VBAプロセスとの通信キューおよび入力監視スレッドを構築する。
# ==============================================================================

# [非同期コマンド受付キューの生成]： PowerShellとVBA間のプロセス間通信(標準入力)をスレッドセーフに受け渡すためのキューを生成する。
$cmdQueue = [System.Collections.Concurrent.ConcurrentQueue[string]]::new()

# [非同期入力監視用Runspaceの作成]： 別スレッドでパイプ入力を監視するためのRunspace（実行空間）を生成する。
$runspace = [runspacefactory]::CreateRunspace()
$runspace.Open()

# [標準入力監視スレッドの開始]： パイプ入力を非同期で監視し、キューへ登録し続ける。
$psThread = [powershell]::Create().AddScript({
    param ($Queue)

    # 親プロセスからの標準入力の常時監視： パイプからの入力を非同期キューへ登録し続ける。
    while ($true) {
        # 標準入力（VBAからのパイプ出力）を1行読み込む。
        $line = [Console]::ReadLine()
        # 空行の場合は処理をスキップし、次の読み込みへ移行する。
        if ([string]::IsNullOrWhiteSpace($line)) { continue }

        # 読み込んだ文字列(JSON)を共有キューへ追加する。
        $Queue.Enqueue($line)
        Start-Sleep -Milliseconds 50
    }
}).AddArgument($cmdQueue)

$psThread.Runspace = $runspace
$psThread.BeginInvoke() | Out-Null

# [親プロセスの特定]： PID再利用やゾンビ化防止のため、親プロセスの情報を記録する。
try {
    $global:ParentPid = (Get-CimInstance Win32_Process -Filter "ProcessId = $PID").ParentProcessId
    $parentName = (Get-Process -Id $global:ParentPid -ErrorAction SilentlyContinue).Name
    Write-DebugLog -Message "[System] 情報: 親プロセス監視開始 (PID: $global:ParentPid, Name: $parentName)" -Level Info
} catch {
    $global:ParentPid = 0
    Write-DebugLog -Message "[System] 警告: 親プロセスの特定失敗" -Level Warning
}

$watchdogSw = [System.Diagnostics.Stopwatch]::StartNew()

# 外部プロセスへREADY通知を送信する。
try { [Console]::WriteLine("READY") } catch {}
#● Write-DebugLog -Message "[System] 成功: エンジン待機状態" -Level Success

# ==============================================================================
# [メインイベントループ]： キューからJSONコマンドを取り出し、対象の関数を動的に実行する。
# ==============================================================================
try {
    while ($global:BrowserForm.Visible) {
        # UIフリーズ防止のためWinFormsイベントを処理する。
        [System.Windows.Forms.Application]::DoEvents()

        $json = $null

        # ゾンビ化を防止する親プロセスの生存監視： PID再利用対策を強化し、親プロセス消失時に道連れ終了を実行する。
        if ($global:ParentPid -gt 0 -and $watchdogSw.Elapsed.TotalSeconds -gt 2) {
            $watchdogSw.Restart()
            
            $parentProc = Get-Process -Id $global:ParentPid -ErrorAction SilentlyContinue
            if (-not $parentProc -or $parentProc.Name -ne $parentName) {
                Write-DebugLog -Message "[System] 警告: 親プロセスの消失(またはPID再利用)検知による道連れ終了の実行" -Level Warning
                break
            }
        }

        # JSONコマンドの取り出しおよび実行： キューから最も古いJSONコマンドを取り出す。
        if ($cmdQueue.TryDequeue([ref]$json)) {
            try {
                # JSONからオブジェクトへの変換：   受信したJSON文字列をPowerShellのカスタムオブジェクトへデコードする。
                $requestObj = $json | ConvertFrom-Json -ErrorAction Stop
                # 実行対象となるコマンド（関数）名を取得する。
                $method = $requestObj.Command

                # QuitまたはExitコマンドにより即時終了する。
                if ($method -in @("Quit", "QUIT", "Exit")) { break }

                # パラメータ付きデバッグログを出力する。
                $paramStr = if ($requestObj.Parameters) { $requestObj.Parameters | ConvertTo-Json -Compress } else { "なし" }
                if ($global:IsDebugMode) {
                    Write-DebugLog -Message "[Engine] 実行: $method | Params: $paramStr" -Level Info
                } else {
                    Write-DebugLog -Message "[Engine] 実行: $method" -Level Info
                }

                # パラメータ辞書を構築する。
                $parameters = @{}
                # パラメータが存在する場合、関数に渡すための連想配列を構築する。
                if ($null -ne $requestObj.Parameters) {
                    # JSONのキーと値を連想配列へマッピングする。
                    foreach ($prop in $requestObj.Parameters.psobject.Properties) {
                        $parameters[$prop.Name] = $prop.Value
                    }
                }

                # コマンドを動的に実行する。
                $res = & $method @parameters | Out-String

                # 実行結果を出力する。
                if (-not [string]::IsNullOrWhiteSpace($res)) {
                    [Console]::WriteLine("[RESULT]$($res.Trim())")
                }
                [Console]::WriteLine("[SUCCESS]")

            } catch {
                # エラー結果を出力する。
                $errMsg = $_.Exception.Message -replace "`r`n", " " -replace "`n", " "
                [Console]::WriteLine("[ERROR]$errMsg")
<# **保留**  
                # ログファイルに詳細なスタックトレース(PositionMessage)を記録する
                $posMsg = if ($_.InvocationInfo) {$_.InvocationInfo.PositionMessage -replace "`r`n", " " -replace "`n", " " } else { "" }
                Write-DebugLog -Message "[Engine] コマンド実行エラー: $errMsg | 発生位置: $posMsg" -Level Error
#>
                # ログファイルへのスタックトレース記録： 波線等の装飾を含まないスタックトレースを記録する。
                $stackTrace = if ($_.ScriptStackTrace) {$_.ScriptStackTrace -replace "`r`n", " <- " -replace "`n", " <- "
                } elseif ($_.Exception -and $_.Exception.StackTrace) {
                    # ScriptStackTraceが空の場合のフォールバック
                    $_.Exception.StackTrace -replace "`r`n", " <- " -replace "`n", " <- "
                } else {
                    "不明"
                }
                Write-DebugLog -Message "[Engine] コマンド実行エラー: $errMsg | トレース: $stackTrace" -Level Error
            }
        }
        Start-Sleep -Milliseconds 20
    }

} finally {
    # ==============================================================================
    # [安全な終了処理（クリーンアップ）]： フォームやプロセスを破棄し、メモリリソースを解放する。
    # ==============================================================================
    Write-DebugLog -Message "[System] 情報: エンジンシャットダウン" -Level Info

    try { $runspace.Close() } catch {}
    try { if ($null -ne $global:BrowserForm) {$global:BrowserForm.Dispose() } } catch {}

    Start-Sleep -Milliseconds 300

    # ゾンビ化防止のため自プロセスを強制終了する。
    Stop-Process -Id $PID -Force
}
