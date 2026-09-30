

<h4 align="center">汎用RPA操作エンジン (VBA × PowerShell × WebView2 Hybrid-Engine)</h4>

🚀 **概要** (Overview)<br>
レガシーな業務システムから複雑なDOM構造（Shadow DOM、多段iframe、動的生成UIなど）を持つモダンWebアプリケーションまで、操作するための汎用ハイブリッドRPAエンジンです。<br>
従来のSelenium等の外部ドライバー依存や、専用クライアントソフトのインストールハードルを低減し、Windows標準の PowerShell 5.1 と WebView2ランタイム のみで稼働する「ゼロ・インストール稼働」を実現しています。<br>

✨ **主要な設計思想** (Core Concepts)<br>
* ゼロ・インストール稼働 (Zero-Install Operation)
外部ライブラリに依存せず、ファイルの配置のみで即時稼働。厳しいセキュリティ要件下でのエンタープライズ運用が可能です。
* ハイブリッド制御による突破力 (Native + CDP + UIA)
JavaScriptによるDOM操作が弾かれる強固なWebシステムに対し、CDP（Chrome DevTools Protocol）やUIAutomation / Win32 APIを用いたOSレベルの物理操作（マウスクリック・キーボード入力）へ自動的にフォールバックします。
* 「基幹」と「拡張」の完全分離 (Core & Extension Architecture)
システム稼働の生命線である「基幹モジュール」と、用途に応じた「拡張モジュール」を分離。Win32 APIやアセンブリのロードを司令塔に一元化し、堅牢なアーキテクチャを構築しています。
* 型と状態の完全維持 (Strict Type Retention)
VBA ⇔ PowerShell ⇔ JavaScript の異言語間プロセス通信において、通信プロトコルを完全なJSONペイロードに規格化。「型の消失」を防ぎ、純粋なデータ型による堅牢なエラーハンドリングを実現しています。

🏗️ **アーキテクチャ図解** (Architecture)<br>
プロセス間の独立性を保ちながら、標準入出力（パイプ）を介して非同期・疎結合に同期をとります。

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

🎯 **主な機能** (Highlights)<br>
🛡️ Robust DOM (レスポンシブ非表示罠の完全回避)<br>
Webサイトで多用される「透明化された入力要素」や「隠蔽されたスマホ用メニュー」の誤爆を防ぎます。
* フェーズ1: 曖昧なXPath指定から、画面に見えている要素だけを抽出し、一意のCSSセレクタを動的生成。
* フェーズ2: 単純な display: none だけでなく、CSS透明度や要素サイズを総合評価し、確実にスクロール＆可視化チェックをしてからクリックを発火。

👻 多段iframe & Shadow DOMの完全透過<br>
WebView2の初期化フェーズにて、Shadow DOMを貫通する再帰探索関数や多段iframeを透過する探索関数をOSレベルで自動注入。これにより、複雑にネストされたモダンコンポーネントにもシームレスにアクセス可能です。

⚙️ OS・ブラウザ間の座標統合メカニズム (DevTools ピッカー)<br>
* マウスポインタ下の要素を自動取得するピッカー機能を搭載。
* OS座標の取得: UIAから要素の中心物理座標を取得。
* 座標変換とDPR補正: ブラウザウィンドウ内での相対物理座標へ変換し、WindowsのDPR（ディスプレイ拡大率）とスクロール量を補正。
* DOMとの照合: CDPで取得した DOMSnapshot の正確な BoundingBox と照合し、最も適した要素を特定してRPAコマンドを自動生成します。

📦 Fetch APIによるサイレントダウンロード<br>
EdgeネイティブのPDFビューア画面などで不安定要因となるOSダイアログを排除。現在のセッションを利用し、Fetch APIでバイナリデータを裏側から直接取得するバックグラウンド保存を実現しています。

📁 **ディレクトリ構成** (Directory Structure)

```text
 [RPA実行ルートフォルダ]
 ├─ Ps_Engine_Core_v-.ps1         (司令塔・ルーター・Win32API/UIA一括ロード)
 │
 ├─ [基幹モジュール] (ブラウザ起動と通信に必須の4ファイル)
 │   ├─ Lib-WebJS_v-.ps1          (ブラウザ全画面へのJS自動注入ユーティリティ)
 │   ├─ Lib-WebView2_Init_v-.ps1  (ブラウザ画面生成・タブ管理)
 │   ├─ Lib-WebView2_Native_v-.ps1 (Native API通信)
 │   └─ Lib-WebCDP_v-.ps1         (WebSocket・CDP高速通信)
 │
 ├─ [拡張モジュール] (用途に応じた機能群)
 │   ├─ Lib-WebAction_v-.ps1      (標準DOM操作・待機)
 │   ├─ Lib-WebXPath_v-.ps1       (XPath操作)
 │   ├─ Lib-WebSafeAction_v-.ps1  (フェイルセーフ・可視化待機クリック)
 │   ├─ Lib-DesktopUIA_v-.ps1     (OSネイティブ・Win32API物理操作)
 │   ├─ Lib-DevTools_v-.ps1       (ピッカー・セレクタ逆算機構)
 │   └─ Lib-WebDebug_v-.ps1       (DOMスナップショット・証跡保存)
 │
 ├─ 実行用マクロファイル.xlsm       (VBA司令塔)
 │
 ├─ [ Libs ]                      (WebView2関連DLL: Core, WinForms, Loader)
 │  ├─ Microsoft.Web.WebView2.Core.dll
 │  ├─ Microsoft.Web.WebView2.WinForms.dll
 │  └─ WebView2Loader.dll
 │
 ├─ [ Logs ]                      (実行ログ・スクショ等の出力先) 5世代管理
 └─ [ UserData ]                  (WebView2独立ユーザーデータフォルダ)
```
 
🛠️ **必要なもの** (Requirements)<br>
* Windows 10 / 11
* PowerShell 5.1 (Windows標準)
* Microsoft Edge WebView2 ランタイム
* Microsoft Excel (VBA環境)
* VBA-JSON v2.3.1 (JsonConverter) _※要 Microsoft Scripting Runtime 参照設定_

📚 **ドキュメント** (Documentation)<br>
モジュール別の詳細な関数リファレンスや、プロセス間通信のJSONプロトコル仕様については、マニュアル（仕様書）をご参照ください。

&emsp;&emsp; _**m01_セットアップ(Getting Started).md**_<br>
&emsp;&emsp; _**m02_rpa-Engine(開発・運用)仕様書.md**_<br>
&emsp;&emsp;&emsp; _m03_Engine(関数TEST)ローカルhtml.md_<br>
&emsp;&emsp;&emsp; _m04_Engine(関数TEST)公開サイト.md_<br>
&emsp;&emsp;&emsp; _m11_コードテスト(ローカルhtml).md_<br>
&emsp;&emsp;&emsp; _m12_コードテスト(公開サイト).md_<br>

---

📜 **ライセンス** (License)

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.<br>
このプロジェクトは **MIT License** の下で公開されています。詳しくは [LICENSE](LICENSE) ファイルをご覧ください。

<br>
<p align="right">
  作成日: 2026年9月27日<br>
  作成者: System beginner & AI Collaborator (Gemini)
</p>

