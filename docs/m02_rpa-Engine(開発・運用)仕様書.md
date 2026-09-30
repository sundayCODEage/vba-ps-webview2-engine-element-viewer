
<h4 align="center">汎用RPA操作エンジン 開発・運用仕様書</h4>
<p align="center"><b>【VBA × PowerShell × WebView2 ハイブリッド制御基盤】</b> (<i>++ たぶん正しい ++</i>)</p>

<p align="right">
  版数: 第2.1版<br>
  最終更新日: 2026年9月27日<br>
  作成者: System beginner & AI Collaborator ( <b>Gemini</b> )
</p>

#### 第1章 システム概要と設計思想
本システムは、レガシーな業務システムから複雑なDOM構造（Shadow DOM、多段iframe、動的生成UIなど）を持つモダンWebアプリケーションまで、高速かつ安全に操作する汎用RPAエンジンです。<br>
従来の自動化ツールが抱えていた「環境構築のハードル」や「操作の不安定さ」を解決するため、以下のコア思想に基づいて設計しています。

#### 1.1 コアとなる設計思想
* **ゼロ・インストール稼働 (Zero-Install Operation)**
外部ライブラリ（Selenium等）や専用のRPAクライアントソフトへの依存を排除しています。Windows OSに標準搭載されている **PowerShell 5.1** と **Microsoft Edge WebView2 ランタイム** のみで駆動するため、ファイルの配置のみで稼働し、厳しいセキュリティ要件下でのエンタープライズ運用を可能にします。
* **「基幹」と「拡張」の完全分離 (Core & Extension Architecture)**
システム稼働の生命線である「基幹モジュール」と、用途に応じて読み込む「拡張モジュール」にアーキテクチャを分離しました。 Win32 APIやアセンブリのロードは司令塔（Ps_Engine_Core）に一元化し、二重ロードによるクラッシュを防ぐ堅牢な疎結合を実現しています。
* **ハイブリッド制御による突破力 (Native + CDP + UIA)**
通常のDOM API（JavaScript）だけでは反応しない厳格なWebシステムに対し、CDP（Chrome DevTools Protocol）を併用した操作や、UIAutomation / Win32 API を用いたOSレベルの物理操作（マウスクリック・キーボード入力）へ自動的にフォールバックする多段構えの突破力を備えています。
* **型と状態の完全維持 (Strict Type Retention)**
言語間（VBA ⇔ PowerShell ⇔ JavaScript）のプロセス通信で生じがちな「型の消失」を防ぐため、通信プロトコルを完全なJSONペイロードに規格化しています。曖昧な文字列判定に依存しない、純粋なデータ型による堅牢なエラーハンドリングを構築しています。

#### 1.2 アーキテクチャ図解
プロセス間の独立性を保ちながら、標準入出力（パイプ）を介して同期をとる非同期・疎結合なアーキテクチャです。

```text
  +-------------------------------------------------------------+
  |                         VBA (Excel)                         |
  |  - JSONペイロード送信 ／ 応答待ち・タイムアウト監視             |
  +------------------------------+------------------------------+
                                 | (標準入出力パイプ通信)
                                 v
  +-------------------------------------------------------------+
  |               PowerShell 5.1 (Ps_Engine_Core)               |
  |  - Win32 API一括ロード ／ JSONルーティング / 例外パース        |
  +------------------------------+------------------------------+
                   (ドットソース結合モジュール群)
              | (基幹モジュール)                 | (拡張モジュール)
              v                                  v
  +-----------------------+              +----------------------+
  | WebView2_Init/Native  |              | WebAction / WebXPath |
  | WebCDP / WebJS        |              | DesktopUIA / DevTools|
  +-----------------------+              +----------------------+
                 WebView2 フォーム・ブラウザコア
              |                                  |
              | (経路A: Native制御)               | (経路B: CDP / UIA制御)
              | - ExecuteScriptAsync             | - WebSocket通信 / Win32 API
              v                                  v
  +-------------------------------------------------------------+
  |    ターゲットWebページ (DOM / iframe / Shadow DOM) / OS       |
  +-------------------------------------------------------------+
```

---

#### 第2章 動作環境とデータ管理
#### 2.1 要求スペック・必須要件

*  Framework: PowerShell 5.1 (Windows OS標準)
*  ランタイム: Microsoft Edge WebView2 ランタイム
*  VBA依存関係: VBA-JSON v2.3.1 (JsonConverter)

> [!NOTE]
プロセス間通信におけるJSON解析のため、VBE（Visual Basic Editor）の参照設定にて Microsoft Scripting Runtime の有効化が必須です。
&emsp;&emsp;`[ツール]` -> [参照設定] -> Microsoft Scripting Runtime に✅<sub>(チェック)</sub>を入れてください。

#### 2.2 標準ディレクトリ構成とモジュール配置
システム一式を任意のローカルフォルダに配置するだけで稼働します。

```text
 [RPA実行ルートフォルダ]
 ├─ Ps_Engine_Core_v-.ps1         (司令塔・ルーター・Win32API/UIA一括ロード)
 │
 ├─ ✨ [基幹モジュール] (ブラウザ起動と通信に必須の4ファイル)
 │   ├─ Lib-WebJS_v-.ps1          (ブラウザ全画面へのJS自動注入ユーティリティ)
 │   ├─ Lib-WebView2_Init_v-.ps1  (ブラウザ画面生成・タブ管理)
 │   ├─ Lib-WebView2_Native_v-.ps1 (Native API通信)
 │   └─ Lib-WebCDP_v-.ps1         (WebSocket・CDP高速通信)
 │
 ├─ ✨ [拡張モジュール] (用途に応じた機能群)
 │   ├─ Lib-WebAction_v-.ps1      (標準DOM操作・待機)
 │   ├─ Lib-WebXPath_v-.ps1       (XPath操作)
 │   ├─ Lib-WebDebug_v-.ps1       (DOMスナップショット・証跡保存)
 │   ├─ Lib-DesktopUIA_v-.ps1     (OSネイティブ・Win32API物理操作)
 │   ├─ Lib-WebSafeAction_v-.ps1  (フェイルセーフ・可視化待機クリック)
 │   └─ Lib-DevTools_v-.ps1       (ピッカー・セレクタ逆算機構)
 │
 ├─ 実行用マクロファイル.xlsm       (VBA司令塔)
 │
 ├─ [ Libs ]                      (WebView2関連DLL: Core, WinForms, Loader)
 │  ├─ Microsoft.Web.WebView2.Core.dll
 │  ├─ Microsoft.Web.WebView2.WinForms.dll
 │  └─ WebView2Loader.dll
 │
 ├─ [ Logs ]                      (実行ログ・スクショ等の出力先) 5世代管理
 │                         (Ps_Engine_Core: $global:CONFIG = @{ MaxLogGenerations = 5 })
 └─ [ UserData ]                  (WebView2独立ユーザーデータフォルダ)
```

#### 2.3 環境分離と画面安定化 (UDFとMaximized)
* **完全な環境分離 (UDF: \\UserData\\EBWebView):**
WebView2はセッションデータやキャッシュ、Cookieを管理するために、実行フォルダ直下に独立したユーザーデータフォルダ（UserData/EBWebView）を生成します。
OS標準のEdgeブラウザやユーザーのWindowsログインセッションとは完全に隔離された「RPA専用のクリーンなブラウザ環境」が構築され、手作業との干渉を防ぎます。アプリケーションを終了してもCookieやLocalStorageは残存し、次回起動時に引き継がれます。
  > **【具体例】** 普段業務でログインしているOffice 365や社内ポータルがあっても、RPAエンジンはそれとは別の「専用のまっさらなブラウザ」として立ち上がります。

* **画面の固定化 (Maximized):**
開発時（ピッカー取得時）と本番実行時における「レスポンシブWebデザインによるDOMの構造変化（スマホビューへの意図せぬ切り替わり等）」を防ぐため、RPAブラウザウィンドウは Ps_Engine_Core でOSの最大有効領域を取得し、常に最大化（Maximized）状態で起動します。これにより物理クリック時の座標ズレ事故もOSレベルで防止します。

#### 2.4 キャッシュのクリアと安定化
長期間の運用時、古いキャッシュがRPAの挙動に悪影響を及ぼすのを防ぐため、テスト開始時やジョブの節目で Clear-WebCache を実行し、クリーンな状態で処理を開始することを推奨します。

```vba
' モードを指定してキャッシュを削除
' CacheOnly: 一時ファイルや画像キャッシュのみ削除
' All: Cookieやローカルストレージを含めた全履歴を完全リセット

rpaEngine.RunAction "Clear-WebCache", CreateParams("Mode", "All")
```

---

#### 第3章 モジュール・コア関数リファレンス
本エンジンは、VBA（司令塔）とPowerShell（実行エンジン）の独立したプロセスで構成され、PowerShell側は機能ごとに分割されたモジュール群を動的に結合（ドットソース読み込み）して稼働します。

#### 3.1 VBA側モジュール（操作指示・司令塔）： 構成例
```text
 ├📊 xxx_rpa.xlsm         # VBA ( シナリオ )
 ├── Ps_Engine.cls       ( プロセス通信とAPI実行を担うRPAエンジンのコアクラス )
 ├── JsonConverter       ( VBA-JSON ) VBAでのJSON解析
 ├── Ps_Bridge.bas       ( JSONパース・エラー変換などのVBA側ユーティリティ )
 │
 ├── フォーム
 └── Mod_DevTools.bas    ( RPAの操作 )
```

* **Ps_Engine (クラスモジュール):**
PowerShellプロセスの非同期起動、JSONペイロードの構築（BuildPayload）、および標準入出力を介した同期通信（SendAndReceive）を担うコアクラス。UIフリーズ防止とクラッシュ検知、ミリ秒単位のタイムアウト監視を内包します。
* **Ps_Bridge (標準モジュール):**
プロセス間通信を補助するユーティリティ群。PowerShellから返された例外文字列を正規表現で解析し、VBAの実行時エラーとして構造化する ParseRpaError などを提供します。
* **Mod_DevTools & フォーム:**
RPAエンジンの操作対象要素を画面上から直接取得し、最適なCSSセレクタやUIA識別子を自動解析してVBAコード（RPAコマンド）を生成するためのGUIピッカー機能と制御ロジック。

#### 3.2 PowerShell コアモジュール群 (Ps_Engine_Core)
* **Ps_Engine_Core (司令塔・ルーティング):** メインループ、JSONルーティング、例外フォーマット生成（New-EngineException）、およびタブ管理を行います。
* **Lib-WebJS (共通ユーティリティ):** ブラウザ初期化時にOSネイティブレベルで自動注入されるグローバルJS関数群（多段iframe透過 utilFindInFrames、可視性判定 isVisible、要素逆引き DOMUtils 等）を定義します。

#### 3.3 ブラウザ制御・通信モジュール
* **Lib-WebView2_Init:** WinFormsの生成、WebView2エンジンの非同期起動、およびJSの自動注入を実施します。ポップアップ（別ウィンドウ）をインターセプトし、仮想タブとして捕獲・ガベージコレクションする機構を持ちます。
* **Lib-WebView2_Native:** WebView2標準の ExecuteScriptAsync を使用したJS実行基盤。JSONアンエスケープと指数バックオフによるリトライ機構を内包します。
* **Lib-WebCDP:** WebSocketセッションの確立とJSON-RPCメッセージの送受信を管理します。DOM操作の高速化と、OSレベルのクリック・キーボード入力（Input.dispatchMouseEvent 等）のエミュレートを行います。

#### 3.4 DOMアクション・UIA連携モジュール
* **Lib-WebAction:** Invoke-WebClick や Set-WebTextInput など、DOMベースの標準操作を提供します。JS操作失敗時には、絶対座標を用いた物理操作（フォールバック）へ自動的に移行します。Fetch APIを用いたサイレントダウンロード機能も内包します。
* **Lib-WebSafeAction:** 透明要素や隠蔽UIの誤爆を防ぐため、可視性の厳密チェックとスクロールを伴う確実なクリック（Invoke-WebSafeClick）を実行します。
* **Lib-WebXPath:** document.evaluate を用いたXPath探索と、表記揺れに対する自動補正（Normalize-XPath）を提供します。
* **Lib-DesktopUIA:** UIAutomationを用いたOSネイティブ操作モジュール。「名前を付けて保存」のOSダイアログ突破（クリップボード経由の安全なパス入力）や、Win32 APIを用いた物理クリック・キー送信を担います。
* **Lib-DevTools (ピッカー):** UIA（OS座標）とWebView2（DOM座標）を統合し、マウスポインタ下の要素から最適なCSSセレクタを動的生成・逆算します。Fuzzy検索（レーベンシュタイン距離）によるフォールバック機構を備えます。
* **Lib-WebDebug:** 画面全体のHTML保存、多段iframeのツリー解析、CDPによるDOMスナップショットのエクスポート機能等を提供します。

---

#### 第4章 プロセス間通信とエラーハンドリング設計
VBA（呼び出し元）とPowerShell（実行エンジン）は独立したプロセスとして稼働するため、標準入出力（パイプ）を介した厳密なメッセージプロトコルで同期をとります。

#### 4.1 JSONペイロードによる通信プロトコル
VBAからPowerShellへのコマンド送信は、必ず以下のJSON構造にシリアライズして標準入力（StdIn）へ送信します。これにより、文字列のエスケープ漏れや引数の順番間違いを防ぎます。

```json
{
  "Command": "Invoke-WebClick",
  "Parameters": {
    "Selector": "button.submit"
  },
  "Settings": {
    "TimeoutSec": 30
  }
}
```

#### 4.2 異常系のパースと伝播
PowerShellからVBAへ状態を返す際は、行頭に必ず以下の識別子を付与し、VBA側はこれを監視し同期処理を行います。

* [RESULT] ...&nbsp; : 関数の戻り値（データ、取得文字列など）。
* [SUCCESS] ..: コマンドが完全に、正常終了したことを示すシグナル。
* [ERROR] ...&nbsp;&nbsp;&nbsp;: PowerShell内で発生した例外メッセージ。

PowerShell側で New-EngineException を用いて生成されたエラー文字列は、VBA側の正規表現パーサー（ParseRpaError）によって RpaExceptionInfo 構造体に変換されます。これにより、「どの関数で」「どのような種別のエラーが」「どのような詳細情報と共に」発生したかが正確にVBAの実行時エラーとして上位プロシージャへ伝達されます。

```vba
ErrorHandler:
    If Err.Source = "PS_Engine" Then
        Dim rpaErr As RpaExceptionInfo
        rpaErr = ParseRpaError(Err.Description)
        
        Select Case rpaErr.ErrorType
            Case "Timeout"
                Debug.Print "【待機超過】画面の応答がありません (" & rpaErr.FunctionName & ")"
            Case "未発見"
                Debug.Print "【スキップ】要素が見つかりません: " & rpaErr.Details
                Resume Next ' エラーを無視して次へ進む運用も可能
            Case "JSエラー"
                MsgBox "システムエラー: " & rpaErr.Message, vbCritical
        End Select
    End If
```

#### 4.3 VBA側の堅牢な3段階監視（SendAndReceive）
プロセスハングアップを防ぐため、VBA側では以下の3つの防波堤を敷いています。

* **プロセスの生存確認:** psProcess.Status を監視し、エンジンの予期せぬクラッシュを即座に検知。
* **日跨ぎ対応タイムアウト:** Windows API GetTickCount を使用し、ミリ秒単位で安全なタイムアウト判定を実施（無限待機の防止）。
* **UIフリーズ防止:** Sleep と DoEvents を組み合わせ、レスポンスを待ちながらもExcel自体の操作性を維持。

#### 4.4 異言語間連携における型統一ラッパーとコーディング規約
PowerShell（制御側）とJavaScript（ブラウザ側）という異なる言語間の通信において発生する「型の消失」や「文字列化による揺らぎ」を完全に排除するため、「JSONラッパー構造の全面的な基準化」を採用しています。

* **純粋な型での Return (JS側):** 成功/失敗の判定は return true; または return false; を使用し、無意味な文字列（return "success"; など）は使用しません。
* **スマートな型評価 (PS側):** 文字列としての泥臭い判定（if ($res -eq "true")）は完全に禁止し、保証された型を直接評価（if ($res) または if (-not $res)）します。
* **末尾のコメント禁止:** JavaScriptをヒアドキュメント（@"..."@）で渡す際、スクリプトの最終行には絶対に // によるコメントを書いてはいけません。ラッパーの閉じカッコ )(); がコメントアウトされ、致命的な構文エラーを引き起こします。

---

#### 第5章 要素の特定と堅牢な操作アプローチ
Web自動化における最大の課題である「要素の待機」と「確実にクリック・入力する技術」に関する仕様です。

#### 5.1 待機処理 (Wait) の実践的な使い分け
Web画面の挙動に合わせて、最適な待機コマンドを選択します。

* **Wait-WebPageLoad (標準・完全待機)**
  * **用途:** 画面遷移時のデフォルト。
  * **特徴:** 画像/CSS等の全リソース読み込み完了と、Ajax通信の沈静化、および全ての子iframeのDOM完成を再帰的にチェックして完全に待ちます。
* **Wait-WebDocumentReady (高速・軽量待機)**
  * **用途:** 背景で重い通信が走り続ける画面や、file:// 環境でのiframe読込制限を回避する際に使用。
  * **特徴:** ルートドキュメントのDOMツリーの構築完了のみ（readyState = 'complete'）で高速に復帰します。
* **Wait-WebScreenUnlock / Wait-WebXPathElementDisappear (マスク待機)**
  * **用途:** 業務システム特有の「処理中」を示す半透明のオーバーレイ（グレーアウト画面）の解除待機。
  * **特徴:** 検索やバッチ起動直後に呼び出すことで、z-indexの高いマスク要素が画面から消滅し、ユーザー操作が可能になるまで確実に待機します。

#### 5.2 CSSセレクタとXPathの指定・自動補正
**CSSセレクタを用いた操作:**（Set-WebTextInput等）は全てiframeを自動探索するため、通常の操作においてフレーム階層を意識する必要はありません。

* **ID指定:** &emsp;&emsp;&nbsp; #login-button
* **クラス指定:** &nbsp; .submit-btn
* **属性指定:** &emsp;&nbsp; input\[name='username']

```vba
' テキストボックスへの入力
rpaEngine.RunAction "Set-WebTextInput", CreateParams("Selector", "input[name='username']", "Value", "user123")

' ボタンのクリック
rpaEngine.RunAction "Invoke-WebClick", CreateParams("Selector", ".submit-btn")
```

**XPathの自動補正 (Normalize-XPath):** 複雑な表構造や特定のテキストを持つ要素を狙う場合はXPathを使用します。
本エンジンでは、//button\[text()='送信'] のような完全一致指定が入力された場合、内部で自動的に \[contains(normalize-space(.), '送信')] に変換（正規化）され、HTMLソース上の余分な改行や空白による「要素が見つからない」エラーを未然に防ぎます。

```vba
' 「承認」というテキストを持つボタンをXPathで狙ってクリック
rpaEngine.RunAction "Invoke-WebXPathClick", CreateParams("XPath", "//button[text()='承認']")
```

#### 5.3 レスポンシブ非表示罠を回避する堅牢な操作 (Robust DOM)
Webサイトで多用される「透明化された本来の入力要素」や、裏側に隠れている「スマホ用の非表示メニュー」などを誤って掴み、クリックが空振りする現象を回避するための安全装置です。

* **フェーズ1: セレクタ逆生成 (Get-WebCssSelectorHint)**
曖昧なXPath指定（例: //a\[contains(., 'ニュース')]）から、画面に見えている要素だけを抽出し、優先度に従って一意のCSSセレクタを動的生成します。
* **フェーズ2: 安全なクリックの実行 (Invoke-WebSafeClick)**
  * **可視性の厳密判定 (isVisible):** 単純な display: none だけでなく、CSSの透明度（opacity !== '0'）や要素のサイズ（幅・高さが0より大きいか）を総合的に評価し、「人間が視覚的に認識し、操作可能か」を厳格に判定します。
  * **安全なクリック (safeClick):** 対象要素を確実に画面中央へスクロール（scrollIntoView({block: 'center'})）させたのち、他の要素（フローティングメニューなど）の裏に潜り込んでいないかを最終確認してからネイティブクリックを発火させます。

#### 5.4 多段iframeとShadow DOMを透過する探索アルゴリズム（グローバル初期化）
複雑な業務システムにおいて、「多段 iframe（フレームの壁）」と「Shadow DOM（カプセル化の壁）」を透過するため、本エンジンはWebView2のネイティブAPIを活用したペイロードの軽量化を実現しています。

**JSペイロードの軽量化とグローバル初期化:**
以前の設計では、コマンド実行のたびに共通関数を送信していましたが、現在はWebView2の初期化フェーズにて CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync() を使用しています。これにより、ページやiframeが生成される瞬間に以下のコア関数群がOSレベルで自動注入されます。

* **deepQuerySelector:** Shadow DOMを貫通する再帰探索関数
* **utilFindInFrames:** &nbsp;&nbsp;&nbsp;&nbsp;多段iframeを透過する再帰探索関数
* **DOMUtils (isVisible, safeClick 等):** 視覚的な判定を行うトラッキングユーティリティ

このアーキテクチャにより、各アクション実行時のデータペイロードが極小化され、CPU負荷の低減とレスポンスの高速化を実現しています。

```powershell
# 【実装例 (Set-WebTextInput の場合)】
function Set-WebTextInput {
    param ($Selector, $Value)
    # 事前にBase64でサニタイズ
    $selB64 = ConvertTo-JsSafeBase64 -Text $Selector
    $valB64 = ConvertTo-JsSafeBase64 -Text $Value
    
    $js = @"
        const selector = decodeURIComponent(escape(atob('$selB64')));
        const val = decodeURIComponent(escape(atob('$valB64')));
        
        // 注入済みの utilFindInFrames / deepQuerySelector を呼び出すだけ
        const found = utilFindInFrames(window, function(win) {
            try {
                const el = deepQuerySelector(selector, win.document);
                if (el) {
                    el.focus();
                    el.value = val;
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
    if (-not $res) { throw "入力に失敗しました: $Selector" }
}
```

---

#### 第6章 OS・ブラウザ間の座標統合メカニズム
本エンジンでは、ピッカーによる要素取得や、JSエラー時の物理操作フォールバックにおいて、「OSの世界」と「ブラウザの世界」を正確に行き来するための高度な座標変換処理を実装しています。

#### 6.1 座標系とDPRの前提知識
* **Screen座標 (OS絶対座標):** モニター全体の左上を起点とする物理ピクセル。マウスの移動や物理クリック（Win32 API）で使用します。
* **Client座標 (ウィンドウ相対座標):** WebView2コントロールの左上を起点とする座標。
* **DPR (Device Pixel Ratio):** Windowsの「ディスプレイの拡大率（125%など）」。これがOSの物理ピクセルとブラウザの論理ピクセルのズレを生む根本原因です。

#### 6.2 パターンA: 【ピッカー実行時】 OS座標からDOM要素を逆引きする (PointReverse)
画面上の要素にマウスを合わせた際、その物理的な位置からブラウザ内のDOM要素を特定する流れ（Get-WebSelectorFromUia 等の処理）です。

1. **OS座標の取得:** UIAから要素の中心の絶対物理座標（Screen X/Y）を取得します。
2. **Client座標への変換:** PointToClient メソッドを使用し、ブラウザウィンドウ内での相対物理座標（Client X/Y）に変換します。
3. **ブラウザ論理座標への変換（DPRとスクロール補正）:** 取得した Client X/Y を DPRで割り算して論理ピクセルに戻し、さらにページのスクロール量（scrollX/Y）を加算します。
$$\\text{uiaCenterX} = \\left(\\frac{\\text{clientX}}{\\text{dpr}}\\right) + \\text{scrollX}$$
4. **DOMとの照合:** この計算結果と、DOMSnapshot で取得済みの BoundingBox 座標を比較し、ユークリッド距離で最も近い要素を特定します。

#### 6.3 パターンB: 【操作フォールバック時】 DOM座標からOS物理クリックを実行する
JavaScriptでの操作（.click() 等）が拒絶される強固な要素に対し、マウスカーソルを移動させて物理クリックを行う流れの処理です。

1. **論理座標の取得:** DOMから対象要素の BoundingBox（絶対論理座標）を取得します。
2. **スクロール量の減算:** 絶対論理座標から現在のスクロール量を引き、画面内での見え方（Client論理座標）を算出します。
3. **物理座標への変換（DPRの乗算）:** 算出した論理座標に DPRを掛け算し、OSが認識できる物理的な Client X/Y に変換します。
4. **Screen座標への変換:** PointToScreen メソッドを使用して、モニター上の絶対物理座標（Screen X/Y）に変換します。
5. **物理クリックの実行:** 算出した Screen X/Y に対して、Win32 API (SetCursorPos と mouse\_event) を用いて強制的にマウスクリック（または SendKeys によるキーボード送信）を発火させます。

---

#### 第7章 高度な自動化アプローチ (Advanced Controls)
#### 7.1 Edge内蔵PDFビューアからの確実なダウンロード機構 (Fetch API)

EdgeネイティブのPDFビューア画面などで発生する「名前を付けて保存」のOSダイアログ操作は不安定要因となります。本エンジンでは物理的な画面操作のアプローチを破棄し、現在のセッション（Cookie等のログイン状態）を利用して Fetch API でバイナリデータを裏側から直接取得する手法を採用しています。

1. 対象のURLへバックグラウンドでリクエストを送り、データを取得する。
2. 取得したデータを「Blob（生のバイナリデータ）」としてメモリに保持する。
3. メモリ上のデータにアクセスするための「一時的な内部URL（ObjectURL）」を作る。
4. 画面には見えない仮想的なアンカータグ（`<a>`）を生成し、`download` 属性を付与して自作自演でクリックする。

```powershell
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
    });
```

これをVBAの Enable-SilentDownload（事前保存先予約）と連動させることで、ダイアログを一切出さずに数ミリ秒でのバックグラウンド保存を実現します。

#### 7.2 仮想タブとガベージコレクション (クラッシュ保護)
業務システムが発行する window.open などのポップアップ要求をOSレベルで捕捉（インターセプト）し、別ウィンドウではなく「エンジン内部の仮想タブ（非表示のWebView2コントロール）」として捕獲・管理します。

親画面との window.opener などの関係性を完全に維持したままバックグラウンドで処理を進め、JS側からの window.close 要求、あるいはプロセスエラー（クラッシュ）を検知した際は、即座にコントロールを破棄してメモリリーク（ゴースト化）を防ぎ、安全に親画面へ復帰する自動クリーンアップ機構を備えています。

#### 7.3 バッチ監視とポーリング制御
Ajax等により非同期で更新されるバッチ処理状態は、再帰的な検索と状態解析で制御します。

1. Wait-WebXPathElementDisappear で画面のオペレーションマスク解除を確実に見届ける。
2. Export-WebTableToCsv でテーブル情報をメモリ内に取得。
3. VBAの Do Until ループ内で目的のステータスに変化するまで繰り返す。

---

#### 第8章 開発・拡張時のコーディング規約
本エンジンを保守・拡張する際は、堅牢性とプロセス間通信の互換性を維持するため、以下のコーディングルールを厳守します。

#### 8.1 命名規則 (PowerShell)
変数スコープと役割を明確にするため、厳密なケース（大文字・小文字）の使い分けを行います。

|項目|適用ルール|例|備考|
|-|-|-|-|
|**パラメータ・グローバル変数**|PascalCase|$TimeoutSec, $global:Tabs|パラメータとして受け取るもの|
|**ローカル変数**|camelCase|$timeoutSec, $xpathEscaped|関数内で加工・定義するもの|
|**時間・間隔変数**|単位を接尾辞化|$WaitTimeMs, $TimeoutSec|Timeout等の曖昧な命名は禁止し、単位（Sec/Ms）を必須とする|
|**VBAからの真偽値引数**|Flg接尾辞|$IsDebugModeFlg|プロセス間通信(文字列渡し)の内部bool値と区別するため|

#### 8.2 JavaScript インジェクションに関する安全基準

PowerShellからJSのヒアドキュメント（@"..."@）に動的な文字列（CSSセレクタやXPath、入力値など）を展開する際、シングルクォート等の単純な文字列置換はXSS的な構文破壊の要因となります。これを防ぐため、以下のカプセル化機構を標準ルールとします。

* **PowerShell側でのエンコード:** 変数を埋め込む前に、必ず ConvertTo-JsSafeBase64 を用いてBase64文字列に変換します。
* **JavaScript側でのデコード:** JSコード内でデコードして使用します。
* **変数のキャメルケース化:** パラメータが PascalCase であっても、JSへ埋め込むための作業変数は必ず camelCase（例: $XPath → $xpathEscaped）に変換してから使用します。

```powershell
# 【標準実装パターン】
$selB64 = ConvertTo-JsSafeBase64 -Text $Selector
$js = @"
    const selector = decodeURIComponent(escape(atob('$selB64')));
    const el = document.querySelector(selector);
    // ...
"@
```

#### 8.3 JavaScript実装時の制約事項
* **変数宣言のモダン化:** JSコード内において var による変数宣言は原則禁止とし、グローバルスコープ汚染を防ぐため再代入が不要な変数は const、必要な変数は let に統一します。
* **JS内での変数命名:** PowerShellのPascalCaseをJSコンテキストに持ち込まず、JavaScriptの標準規約に従い camelCase を徹底します。
* **末尾コメントの禁止:** JSヒアドキュメントの最終行（"@` の直前行）に // によるコメントを記述してはいけません。ラッパーの閉じカッコ )(); がコメントアウトされ、致命的な構文エラーを引き起こします。

```javascript
// 【NGパターンの具体例】
var el = document.querySelector('.btn');
return el !== null; // ここで要素の有無を返す

// ※ 上記コードがラッパーに組み込まれると閉じカッコ )(); まで全てコメントアウトされ、致命的な構文エラーを引き起こします。
```

#### 8.4 型統一ラッパー構造とプロセス間通信の運用ルール
JS側の純粋なデータ型（Boolean, Number, Null等）をPowerShell側へ正確に引き継ぐため、最下層のJS実行部（Invoke-WebView2NativeScript, Invoke-CdpScript）には以下のラッパーを標準実装しています。

```javascript
// 【JS側ラッパーの構造】
const result = (function() { /\* 動的生成されたJSコード \*/ })();
return JSON.stringify({
    status: "success",
    data: result !== undefined ? result : null
});
```

これにより、過去の資産に存在した $res -eq "true" のような文字列依存の判定コードを完全に排除しています。

**スクリプト実行の共通窓口 (Invoke-WebScript):**
すべてのDOM操作や値の取得、要素の待機処理において、JavaScriptを実行する際の最上位の共通窓口（ルーター）となるのが Invoke-WebScript です。
* **ハイブリッド・ルーティング:** CDPポートが有効な場合は、低レイテンシでネイティブイベントを発火できる Invoke-CdpScript を選択し、無効な場合は Invoke-WebView2NativeScript を選択します。
* **フェイルセーフ:** CDP通信エラー検知時は、システムを停止させることなくNative側へ自動的に処理を切り替えてリトライを実行する堅牢な二重化構造を組み込みました。

**【PowerShell側の運用ルール】**
以降の開発においても、戻り値は「純粋な型」として評価してください。

* **泥臭い文字列判定の禁止**: 過去の資産に見られる $res -eq "true" のような、文字列としての保険的な比較は書かない。
* **スマートな型評価**: 戻り値は完全に型が保証されているため、以下のように直接評価します。

```powershell
$res = Invoke-WebScript -Js $js

# 成功判定 (OK)
if ($res) { return $true }
# 失敗判定 (OK)
if (-not $res) { throw "要素が見つかりません" }

# 文字列比較 (NG: 絶対にやらないこと)
if ($res -eq "true") { ... }
```

---

#### 第9章 デバッグおよび証跡エクスポート
テストエラー時や監査用エビデンスとして、画面の状態をファイル出力が可能です（保存先は実行時セッションごとのLogsフォルダ内で世代管理されます）。

* **Export-WebHtml:** 画面全体（多段iframe含む）のHTMLを保存。
* **Export-WebScreenshot:** 画面のPNGスクリーンショットを保存。
* **Export-WebElementsToCsv:** 画面内の操作可能な全要素の属性値と可視状態をCSV化。
* **Export-WebFrameTreeToCsv:** iframe/frameのネスト階層構造をツリー形式でCSV化。
* **Export-WindowHierarchyToCsv:** OS上のウィンドウ階層（プロセス・ハンドル）のCSV化。
* **Export-WebDomSnapshot:** CDPを利用し、画面のDOMツリー構造とスタイル情報をJSON形式でエクスポート。

---

#### 第10章 要素ピッカーの自動生成メカニズムと処理フロー (DevTools)
画面上の要素をマウスポインタで指定し、最適なRPAコマンドを自動生成する「ピッカー機能」は、VBAからの呼び出し窓口となる主要関数を起点とし、多数の内部関数が連携（オーケストレーション）することで動作します。

#### 10.1 ピッカー処理の全体フロー図
処理は大きく分けて「事前情報の取得」「UIA要素の取得」「ルートごとの探索」「コード組み立て」の4フェーズで進行します。

```text
 [VBAからの呼び出し窓口]
 ├─ Invoke-UiaRecordStep     (カーソル位置から自動取得)
 └─ Test-WebPointReverse     (指定座標から直接取得検証)
      │
      ▽
 [フェーズ1: 事前情報の取得 (DOMスナップショット)]
 └─ Get-DomSnapshotWithBoxModel   (CDP経由で全要素の座標とスタイルを一括取得 ※Lib-WebDebug)
      │
      ▽
 [フェーズ2: UIA要素と座標の取得]
 ├─ Get-UiaTargetInfoFromCursor   (最前面要素と適正な親コンテナの探索)
 ├─ Get-UiaElementFromPoint       (物理座標からのUIA要素取得)
 └─ Get-UiaWindowFromElement      (所属するウィンドウ情報の特定)
      │
      ▽
 [フェーズ3: オーケストレーターによる探索 (Process-RecordAction)]
 ├─ ルートA: 座標ベースのDOM逆引き (優先)
 │    ├─ Invoke-BrowserPointReverse
 │    └─ Get-CorePointReverseJs (js: elementFromPoint等を利用)
 │
 ├─ ルートB: UIAメタデータによるDOM照合 (代替)
 │    ├─ Invoke-UiaDomResolution
 │    ├─ Get-WebSelectorFromUia       (位置・テキストによる厳密マッチ)
 │    └─ Get-WebSelectorFromUiaFuzzy  (レーベンシュタイン距離を用いた曖昧検索)
 │
 └─ ルートC: デスクトップ物理操作へのフォールバック (最終手段)
      └─ Build-FallbackStepCode       (OSダイアログ等、ブラウザ外の場合)
      │
      ▽
 [フェーズ4: RPAコマンドの最終組み立て]
 ├─ Build-RpaStepCode            (操作対象のメタデータ構築)
 └─ ConvertTo-RpaCommandCode     (VBAで実行可能な Invoke-WebClick 等のJSON文字列出力)
```

#### 10.2 メイン処理フロー
1. **エントリーポイント (VBA連携)**
   * **Invoke-UiaRecordStep:** ユーザーの待機秒数を管理し、操作前と操作後の「DOM Snapshot（全要素の座標とスタイル情報）」をCDP経由で一括取得してから、カーソル下の要素解析を開始します。
   * **Get-DomSnapshotWithBoxModel:** 画面操作の直前に、CDP経由でブラウザ画面内の操作可能な全要素の「DOMツリー構造」と「BoundingBox（絶対座標）」を一括取得し、メモリ上に一時保存（スナップショット）します。これが後続の照合処理におけるマスターデータとなります。
   * **Test-WebPointReverse:** VBAから直接「X/Y座標」を渡し、その位置にある要素が正確に取得できるかをテスト（診断）するための主要関数です。

2. **ターゲット要素の特定** (**Get-UiaTargetInfoFromCursor**)
現在のマウスカーソルのOS物理座標（X/Y）から、その下にある最前面のUIA要素を取得します。WindowやPaneなどの巨大なコンテナを避け、クリックや入力に最適な親要素をヒューリスティック（経験則）に探索・スコアリングして特定します。

3. **コマンド生成オーケストレーター** (**Process-RecordAction**)
特定されたUIA要素の情報をもとに、以下のルートでDOM要素への逆引きとコマンド生成を行います。
   * **ルートA (物理座標ベース逆引き):** Invoke-BrowserPointReverse を呼び出し、OSの物理座標をブラウザの論理座標へ変換します。Get-CorePointReverseJs によって構築されたJS (elementFromPoint) を用いてDOM要素を特定し、CDPで取得済みの BoundingBox（絶対座標）と照合・補正します。
   * **ルートB (UIAメタデータベース逆引き):** 座標から特定できなかった場合、Invoke-UiaDomResolution へ移行し、UIAの Name や AutomationId などのテキスト情報をもとにDOM要素を特定します。
   * **ルートC (デスクトップ操作へのフォールバック):** 対象要素がWebブラウザ外（OSの「名前を付けて保存」ダイアログなど）であった場合、Build-FallbackStepCode を用いて、Lib-DesktopUIA 連携用の Win32 API 物理操作コマンドを生成します。

4. **コマンド文字列の最終組み立て** (**ConvertTo-RpaCommandCode**)
特定されたCSSセレクタ、座標（BoundingBox）、およびHTML構造をもとに、VBAから直接呼び出し可能なRPAコマンド（例: Invoke-WebClick -Selector ...）をJSON形式の文字列として最終出力し、VBA側へ返却します。

#### 10.3 セレクタ安定性テスト (Stability Test)： 重要
ピッカーで要素を取得した直後、エンジン内部では Test-SelectorStability による「セレクタの安定性テスト」が自動実行されます。一見すると同じ要素を複数回探すだけの処理に見えますが、これはWeb自動化において重要な防波堤です。

**モダンWeb特有の罠と回避メカニズム:**

* **動的IDの排除:** 業務システムには、アクセスやマウスホバーのたびにIDが変わる要素（例: id="ext-gen-1024"）が頻繁に存在します。一瞬だけ取得したIDを信じて自動化コードを生成すると、本番実行時に要素が見つからないエラー（Flakyテスト）の温床となります。
* **仮想DOMの再描画対策:** ReactやVueなどのモダンフレームワークでは、画面のローディング直後やイベント発生時に、見た目は同じでも裏側のHTML要素が別のものにすり替わっていることがあります。

  > 安定性テストは、取得したセレクタが「数ミリ秒～数秒後も生存し続けているか（一過性の幻ではないか）」を連続検証します。この事前検証を通過したセレクタのみを最終的なRPAコマンドとして採用することで、本番環境におけるエンジンの「止まらない・壊れない」堅牢性を担保します。

---

#### 第11章 モジュール・コア関数リファレンス完全版
本エンジンに実装されている、モジュール別の全関数リファレンスです。（※ `[内部関数]` と記載のあるものはシステム内で自動的に呼び出される裏側処理です）

#### 11.1 PowerShell コアモジュール群
#### 🛠️ [Core] 司令塔・ルーティングモジュール
* **Write-DebugLog:** コンソール出力とファイル出力（世代管理対応）を行うロギング機能。
* **New-EngineException:** **[ERROR]** プレフィックスでVBAへ返す例外文字列をフォーマット生成。
* **Get-ActiveWebView:** 現在アクティブなタブのWebView2インスタンスを取得。
* **Set-ActiveTab:** タブIDを直接指定してアクティブタブを切り替え（前面化）。
* **List-Tabs:** 起動中の全タブ情報（ID、URL、タイトル）をJSONで取得。
* **Switch-Tab:** タブIDによる切り替え。CDPの再接続処理も包含。
* **Switch-TabByTitle:** タイトルの部分一致検索によるタブ切り替え。
* **Wait-Condition:** UIフリーズを防止しつつ、指定条件がTrueになるまで待機（汎用）。
* **Invoke-WebScript:** JS実行のルーティング。CDPが有効ならCDP、失敗時はNativeへフォールバック。
* **Set-EngineConfig:** 実行時のエンジン設定（要素ハイライトのON/OFF等）を動的に変更。

#### 🚀 [Init] & [Native] ブラウザ初期化・ネイティブ通信
* **Clear-WebCache:** UDFのキャッシュ、Cookie、LocalStorage等を非同期で完全削除。
* **Invoke-WebView2NativeScript:** **ExecuteScriptAsync** を使用したJS実行。JSONアンエスケープとリトライ機構を内包。

#### ⚡ [CDP] 高速通信モジュール (WebSocket)
* **Connect-CdpSession:** **json** エンドポイントからTargetIdを探査し、WebSocketセッションを確立。`[内部関数]`
* **Invoke-CdpCommand:** JSON-RPCメッセージの送受信。タイムアウトと自動再接続を管理。
* **Invoke-CdpScript:** CDP経由でのJS評価 (**Runtime.evaluate**)。戻り値のJSONデコードを含む。
* **Invoke-CdpNativeClick:** CDPを使用し、OSレベルのマウスダウン/アップイベントを座標指定でエミュレート。
* **Set-CdpNativeTextInput:** CDPを使用し、キーボード入力をOSレベルでエミュレート（SPA対策）。

#### 11.2 DOMアクション・UIA連携・拡張モジュール群
#### 🛡️ [Action] Web標準操作モジュール
* **Invoke-WebNavigation:** 指定URLへのページ遷移を実行。
* **Wait-WebPageLoad:** DOMの **readyState=complete** を全iframe含めて再帰的に待機。
* **Wait-WebDocumentReady:** 画面全体の読み込みステータス完了を待機。
* **Wait-WebUrlContains**  **Wait-WebTitleContains:** URLやタイトルに指定文字列が含まれるまで待機。
* **Wait-WebElement:** 指定要素がDOM上に出現し、かつ画面上に可視化されるまで待機。
* **Wait-WebElementInFrame:** 指定したiframe内の要素が出現・可視化されるまで待機。
* **Invoke-WebClickInFrame:** 指定したiframe内の要素をスクロールしてクリック。
* **Wait-WebElementInvisible:** 指定要素が非表示になる、またはDOMから消滅するまで待機。
* **Wait-WebScreenUnlock:** 業務システム特有のローディングマスク（透過レイヤー）の解除を待機。
* **Invoke-WebClick:** 多段iframeを透過的に探索し、対象要素をクリック。
* **Set-WebTextInput:** テキストボックスに値を入力し、**input** / **change** イベントを発火。
* **Select-WebDropdown:** ドロップダウン（**select**）の指定値を選択し、**change** イベントを発火。
* **Set-WebCheckbox:** チェックボックスの状態（True/False）を判定し、差異があれば切り替え。
* **Get-WebText:** 要素の **innerText** または **value** を取得。
* **Get-WebUrl** / **Get-WebTitle:** 現在のURL、およびページタイトルを取得。
* **Enable-SilentDownload:** DLダイアログを抑制し、指定フォルダ・ファイル名での裏側ダウンロードを有効化。
* **Wait-FileDownload:** **.crdownload** の消失および排他ロック解除を確認し、DL完了を待機。
* **Get-WebEmbedPdfUrl:** 埋め込みPDF（**embed**）の絶対URLを安全に取得。
* **Test-WebElemente:** Wait-WebElement系の指定要素の存在確認を行い、例外を投げず True/False の文字列を返す。
* **Get-WebAttribute:** 指定したWeb要素の特定の属性値(Attribute)を取得を取得する。
* **Invoke-WebFetchDownload:** Fetch APIを利用した裏側でのサイレントダウンロード発火。
* **Convert-BoundingBoxToScreenPoint:** BoundingBox(絶対論理座標)をOSの絶対スクリーン座標へ変換する。`[内部関数]`

#### 🛡️ [SafeAction] フェイルセーフ・安全クリック（Robust DOM）
* **Get-WebCssSelectorHint:** 曖昧なXPathから、可視状態の要素を厳密に判定し、レイアウト変更に強い一意のCSSセレクタを逆生成する。
* **Invoke-WebSafeClick:** 生成されたCSSセレクタを使用し、スクロールと可視性の最終確認を行った上で、隠し要素の誤爆を防ぎ安全にクリックを実行する。

#### 🎯 [XPath] XPath特殊操作モジュール
* **Normalize-XPath:** XPathの表記揺れ（改行・空白）を自動補正。`[内部関数]`
* **Wait-WebXPathElement:** XPath指定で要素の可視化を待機。デバッグ時は赤枠ハイライトを実行。
* **Wait-WebXPathElementDisappear:** XPath要素の非表示・消滅を待機。
* **Invoke-WebXPathClick:** XPath要素に対し、hover/mousedown/up等の一連のマウスイベントを完全エミュレート。
* **Set-WebXPathTextInput:** XPath要素へフォーカスし、テキスト入力と各種イベント発火を実行。
* **Get-WebXPathText:** XPath要素のタグを判別し、適切なテキスト（**value** または **innerText**）を取得。

#### 🖥️ [UIA] デスクトップ操作モジュール
* **Switch-AppWindow:** Win32 APIを用いて指定した外部ウィンドウを最前面へ引き上げ。
* **Invoke-UiaAction:** UIAutomationを用い、バックグラウンドパターンまたは物理キー送信でOS要素を操作。
* **Invoke-UiaSafeSaveAs:** 「名前を付けて保存」ダイアログを捕捉し、クリップボード経由でパスを入力・保存。
* **Invoke-DesktopSendKeys:** 対象のウィンドウへ物理キー（SendKeys）を送信。
* **Invoke-DesktopCenterClick:** アクティブウィンドウの中央を物理クリックする。
* **Invoke-DesktopCoordinateClick:** 指定されたX/Y座標（OS絶対スクリーン座標）に対して物理マウスクリックを実行する。

#### 🐛 [Debug] デバッグ・証跡モジュール
* **Export-WebHtml:** クロスオリジンを考慮し、全iframeを含むHTMLスナップショットを保存。
* **Export-WebScreenshot:** CDP、またはネイティブAPIへフォールバックして画面のPNGスクショを保存。
* **Export-WebTableToCsv:** テーブル要素を解析し、VBA取込用の配列文字列を含むCSVを生成。
* **Export-WebElementsToCsv:** 画面内の操作可能要素（**input**, **a**, **button** 等）の属性を総ざらいしてCSV化。
* **Export-WebFrameTreeToCsv:** 多段iframeのネスト構造をツリー形式で解析しCSV化。
* **Export-WindowScreenshot:** Win32 API等を使用し、ブラウザの枠を含むウィンドウ全体のスクショを保存。
* **Export-WindowHierarchyToCsv:** OS上で起動している全プロセスのハンドルとタイトル一覧をCSV出力。
* **Write-DebugTextFile:** 任意の文字列をデバッグ用テキストファイルへ追記保存。
* **Export-WebDomSnapshot:** 画面上のDOMツリーとスタイルをCDP経由で一括取得し、JSONとして保存。
* **Export-WebLayoutDump:** **Page.getLayoutMetrics** を用い、DPRやViewportを含むレイアウト情報を保存。
* **Test-SelectorStability:** 指定したセレクタが、画面の再描画等に負けず安定して要素を取得できるかを検証。
* **Test-AllDomSelectorsStability:** 取得済みのDOM Snapshot内の全要素に対して、一括で安定性テストを実行。
* **Convert-DomSnapshotFlat:** CDP固有の圧縮された1次元配列を、解析しやすいフラットなオブジェクトリストに展開する。`[内部関数]`
* **Get-DomSnapshotWithBoxModel:** インタラクティブな操作対象要素に限定し、DOM構造とBoxModel(物理座標)を統合取得して通信負荷を軽減する。`[内部関数]`

#### 🔍 [DevTools] 要素ピッカー・コード自動生成モジュール
* **Invoke-UiaRecordStep:** マウスポインタ下の要素を特定し、最適なCSSセレクタとRPAコマンドコードを自動生成する（ピッカー機能のメイン窓口）。
* **Test-WebPointReverse:** (指定)座標からの要素取得検証(テスト)。
* **Test-WebSelectorResolution:** UIA要素からのセレクタ解決能力（Fuzzy検索等へのフォールバック挙動）を検証・ログ出力する診断関数。
* Process-RecordAction: 取得した要素情報をもとに、DOM逆引き・UIA照合・フォールバック等のルート分岐を制御する。`[内部関数]`。 (他 Get-WebSelectorFromUia 、、、)
