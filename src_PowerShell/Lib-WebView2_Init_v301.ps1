# ------------------------------------------------------------------------------#●#✘ 
<#
  [WebView2 初期化モジュール]： RPA用ブラウザウィンドウの生成、WebView2エンジンの起動、共通JSの自動注入、及び
  ポップアップ（新規ウィンドウ）のタブ化管理とクラッシュ保護を行う。
#>
# ------------------------------------------------------------------------------

# [ブラウザシステムの初期化]： 必須DLLの動的ロードとベースとなるウィンドウ生成を開始する。
Write-DebugLog -Message "[System] 開始: ブラウザシステムの初期化 ..." -Level Info
$libDir = Join-Path $global:ScriptDirectory "Libs"

# [必須WebView2拡張DLLの動的読み込み]： CoreおよびWinFormsのDLLをロードし、WebView2の実行環境を確立する。
$requiredDlls = @(
    "Microsoft.Web.WebView2.Core.dll",
    "Microsoft.Web.WebView2.WinForms.dll"
)

foreach ($dllName in $requiredDlls) {
    $dllPath = Join-Path $libDir $dllName
    if (Test-Path $dllPath) {
        try {
            $versionInfo = (Get-Item $dllPath).VersionInfo.FileVersion
            Add-Type -Path $dllPath
            Write-DebugLog -Message "[System] 成功: DLLロード ($dllName Version: $versionInfo)" -Level Success
        } catch {
            Write-DebugLog -Message "[System] 致命的エラー: DLLロード失敗 ($dllName - $($_.Exception.Message))" -Level Fatal
            exit 1
        }
    } else {
        Write-DebugLog -Message "[System] 致命的エラー: DLLが存在しません ($dllName)" -Level Fatal
        exit 1
    }
}

# [RPAブラウザ用ウィンドウの生成]： WinFormsを利用し、WebView2コントロールをホストするための土台となるフォームウィンドウを構築する。
$global:BrowserForm = New-Object System.Windows.Forms.Form -Property @{
    # フォームのタイトルバーにセッションIDを含める。
    Text   = "RPA Browser - $global:AppSessionId"
    # 設定値からウィンドウの幅と高さを指定する。
    Width  = $global:CONFIG.BrowserWidth
    Height = $global:CONFIG.BrowserHeight
    # 画面の中央に配置し、ウィンドウを最大化する。
    StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    WindowState   = [System.Windows.Forms.FormWindowState]::Maximized
}

# [WebView2コントロールの配置]： フォームへの紐付けと画面領域の最大化設定を行う。
# ブラウザ本体となるWebView2コントロールのインスタンスを生成する。
$global:WebViewCtrl = New-Object Microsoft.Web.WebView2.WinForms.WebView2
# コントロールをフォーム全体に広がるようにDock設定する。
$global:WebViewCtrl.Dock = [System.Windows.Forms.DockStyle]::Fill
# フォームのコントロールコレクションへWebView2を追加する。
$global:BrowserForm.Controls.Add($global:WebViewCtrl)
# フォームを画面上に表示する。
$global:BrowserForm.Show()

try {
    # ユーザーデータフォルダの準備： ユーザーデータフォルダが存在しない場合は作成する。
    $userDataFolder = Join-Path $global:ScriptDirectory "UserData"
    if (-not (Test-Path $userDataFolder)) {
        New-Item -ItemType Directory -Path $userDataFolder | Out-Null
    }

    # WebView2起動オプションの設定： 初期化用のオプションクラスを生成する。
    $options = [Microsoft.Web.WebView2.Core.CoreWebView2EnvironmentOptions]::new()

    # ディスクキャッシュの無効化・制限設定： RPA動作の安定化のため、キャッシュ利用を最小限に抑える。
    $cacheArgs = " --disable-disk-cache --disk-cache-size=1 --media-cache-size=1"

    if ($global:CdpPort -gt 0) {
        # CDPポートが指定された場合、リモートデバッグオプションを追加して引数を構成する。
        $options.AdditionalBrowserArguments = "--disable-gpu --disable-dev-shm-usage --remote-debugging-port=$global:CdpPort" + $cacheArgs
        Write-DebugLog -Message "[System] 起動モード: CDP有効 (Port: $global:CdpPort)" -Level Info
    } else {
         # CDPポート指定がない場合、標準的な動作安定化の引数のみを構成する。
        $options.AdditionalBrowserArguments = "--disable-gpu --disable-dev-shm-usage" + $cacheArgs
        Write-DebugLog -Message "[System] 起動モード: 標準 (CDP無効)" -Level Info
    }

    # WebView2エンジンを非同期で起動： オプションを適用してブラウザコアプロセス(Environment)を起動する。
    $envTask = [Microsoft.Web.WebView2.Core.CoreWebView2Environment]::CreateAsync(
        $null, $userDataFolder, $options
    )
    while (-not $envTask.IsCompleted) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 10
    }

    # WebView2コントロールを初期化： 起動したエンジンをWinForms上のコントロール(WebViewCtrl)へ紐付ける。
    $initTask = $global:WebViewCtrl.EnsureCoreWebView2Async($envTask.Result)

    while (-not $initTask.IsCompleted) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 10
    }

    # 共通JSユーティリティのグローバル自動注入： ページやiframeが生成されるたびに、エンジン共通のJS関数をネイティブレベルで自動登録する。
    $combinedJsUtils = "$global:ENGINE_JS_UTILS`n$global:ENGINE_DEBUG_JS_UTILS"
    $scriptTask = $global:WebViewCtrl.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync($combinedJsUtils)
    
    while (-not $scriptTask.IsCompleted) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 10
    }
    Write-DebugLog -Message "[System] 成功: WebView2エンジン初期化およびJS自動注入" -Level Success

    # 初回タブ(P01)の登録： 最初の起動画面を「P01」としてタブ管理辞書へ登録し、以降のポップアップ(仮想タブ)と区別する。
    $global:Tabs["P01"] = @{
        # 実体となるWebView2コントロールを保持する。
        WebView = $global:WebViewCtrl
        # 親を持たないメインタブであることを示すためnullを設定する。
        Parent  = $null
        # 現在のURLを文字列として保持する。
        Url     = $global:WebViewCtrl.Source.ToString()
    }
    # アクティブタブIDをP01に設定する。
    $global:ActiveTabId = "P01"

    # 初回タブ(P01)のURL更新イベント登録： 画面遷移が完了するたびに、内部管理用のURL文字列を同期させる。
    $global:WebViewCtrl.CoreWebView2.add_NavigationCompleted({
        param($sender, $e)
        try {
            $global:Tabs["P01"].Url = $sender.Source
            Write-DebugLog -Message "[P01] URL更新: $($sender.Source)" -Level Info
        } catch {}
    })

    # ランタイムバージョンのログ出力： 接続先となるWebView2のバージョン情報を記録する。
    $runtimeVersion = $global:WebViewCtrl.CoreWebView2.Environment.BrowserVersionString
    Write-DebugLog -Message "[System] 情報: 接続先ランタイム (Version: $runtimeVersion)" -Level Info

    <#
      【状態遷移】ポップアップ（子画面）要求のキャプチャと仮想タブ化
     ------------------------------------------------------------------------------
      (1. 発生: 業務システムのJSが window.open() や target="_blank" を発行する。)
       2. 捕捉: 本イベントでインターセプトし、OS標準の別ウィンドウ起動を完全にブロックする。
       3. 生成: エンジン内部で新しい WebView2 コントロール（仮想タブ）を生成する。
       4. 連携: 親画面との親子関係 (window.opener) を完全に維持したままバックグラウンドでロードさせる。
       5. 破棄: 業務完了後、JS側の window.close() 要求を検知してタブを自動破棄し、親画面へ復帰する。
    #>
    $global:NewWindowHandler = {
        param($sender, $e)
        Write-DebugLog -Message "[WebView2] 情報: 新規ウィンドウ要求を検知。仮想タブとして捕獲します。" -Level Info

        try {
            # (ステップ 2) 捕捉： 非同期でWebView2コントロールを生成するため、要求を一時停止(Deferral)してOS側の処理進行を待たせる。
            $deferral = $e.GetDeferral()

            # (ステップ 3) 生成： 新しい仮想タブとして WebView2 コントロールをバックグラウンドで生成する。
            $newWebView = New-Object Microsoft.Web.WebView2.WinForms.WebView2
            $newWebView.Dock = [System.Windows.Forms.DockStyle]::Fill

            # コールバック関数内で使用する変数を Tag プロパティに退避して引き継ぐ。
            $newWebView.Tag = @{
                Parent   = $global:WebViewCtrl
                Deferral = $deferral
                EventArg = $e
            }

            # WinForms特有のハックを適用する： Handle プロパティにアクセスすることでOSレベルのウィンドウハンドルを強制的に即時生成させ、後続の非同期初期化を安定させる。
            $dummyHandle = $newWebView.Handle

            # WebView2エンジン準備完了時のコールバック処理： 初期化成功時のタブID採番とJS自動注入を行う。
            $newWebView.add_CoreWebView2InitializationCompleted({
                param($senderInit, $argsInit)

                # Tagプロパティに退避しておいた変数群（親コントロール、イベント引数、保留状態）を復元する。
                $state      = $senderInit.Tag
                $parentCtrl = $state.Parent
                $deferral   = $state.Deferral
                $evtArgs    = $state.EventArg

                try {
                    if ($argsInit.IsSuccess) {
                        # 新しい仮想タブに一意のID(GUID)を割り当て、親画面との関係性をグローバル辞書に記録する。
                        $tabId = [guid]::NewGuid().ToString()
                        $global:Tabs[$tabId] = @{
                            WebView = $senderInit
                            Parent  = $parentCtrl
                        }
                        $global:ActiveTabId = $tabId

                        # ポップアップ（仮想タブ）にも共通JSユーティリティを自動注入する。
                        try {
                            $combinedJsUtils = "$global:ENGINE_JS_UTILS`n$global:ENGINE_DEBUG_JS_UTILS"
                            $senderInit.CoreWebView2.AddScriptToExecuteOnDocumentCreatedAsync($combinedJsUtils) | Out-Null
                        } catch {
                            Write-DebugLog -Message "[WebView2] 警告: ポップアップへのJS自動注入に失敗" -Level Warning
                        }

                        try {
                            # CDPを利用して、このポップアップ画面を直接操作・デバッグする際に必要となる内部ID(TargetId)を抽出して保持する。
                            $global:Tabs[$tabId].TargetId = $evtArgs.WindowFeatures.TargetId
                        } catch {
                            Write-DebugLog -Message "[WebView2] 警告: targetId の取得に失敗" -Level Warning
                        }

                        # 生成した仮想タブ(WebView2)を実際の画面に追加し、最前面に表示して親画面を覆い隠す。
                        $global:BrowserForm.Controls.Add($senderInit)
                        $senderInit.BringToFront()

                        #  (ステップ 4) 連携： 捕捉していたポップアップ要求の出力先を、今作成した仮想タブ(NewWindow)に接続する。
                        # - （これにより、JS側の window.opener の参照関係が切断されずに維持される。）
                        $evtArgs.NewWindow = $senderInit.CoreWebView2
                        # OS標準ブラウザや別ウィンドウでの起動を完全にブロックする。
                        $evtArgs.Handled   = $true
                        # 保留していた要求処理を再開する。
                        $deferral.Complete()

                        $global:WebViewCtrl = $senderInit

                        # ポップアップ画面のURL更新イベントを登録する。
                        $senderInit.CoreWebView2.add_NavigationCompleted({
                            param($navSender, $navArgs)
                            try {
                                $global:Tabs[$tabId].Url = $navSender.Source
                                Write-DebugLog -Message "[$tabId] URL更新: $($navSender.Source)" -Level Info
                            } catch {}
                        })

                        # 多段ポップアップ（ポップアップ画面からのさらなるポップアップ）にも対応させる。
                        $senderInit.CoreWebView2.add_NewWindowRequested($global:NewWindowHandler)
                        
                        # (ステップ 5) 破棄： JSからの window.close() 要求に伴う自動クリーンアップ（業務完了やキャンセル等）を行う。
                        $senderInit.CoreWebView2.add_WindowCloseRequested({
                            param($closeSender, $closeArgs)
                            
                            # UIスレッド上で安全にコントロールを操作するため、MethodInvokerデリゲートを生成する。
                            $mi = [System.Windows.Forms.MethodInvoker]{
                                try {
                                    Write-DebugLog -Message "[WebView2] 情報: ポップアップの終了要求(window.close)を検知。タブ($tabId)を破棄します。" -Level Info

                                    # 透明なゴーストとして残るのを防ぐため、コントロールを削除してメモリを解放する。
                                    $global:BrowserForm.Controls.Remove($senderInit)
                                    # WebView2コントロールのメモリやリソースを即座に破棄する。
                                    $senderInit.Dispose()
                                    
                                    # 管理リストから完全削除する。
                                    if ($global:Tabs.ContainsKey($tabId)) {
                                        # タブ管理用のハッシュテーブルから該当のキーと値を消去する。
                                        $global:Tabs.Remove($tabId)
                                    }

                                    # アクティブタブが閉じられたら、自動的に親画面へ戻る。
                                    if ($global:ActiveTabId -eq $tabId) {
                                        # まず、このポップアップを開いた親ウィンドウのTabIdを探す。
                                        $nextTabId = $null
                                        # 登録されている全タブの中から、親コントロールと一致するものを探索する。
                                        foreach ($key in $global:Tabs.Keys) {
                                            if ($global:Tabs[$key].WebView -eq $parentCtrl) {
                                                # 親が見つかった場合は変数へIDを格納し、ループを抜ける。
                                                $nextTabId = $key
                                                break
                                            }
                                        }

                                        # 親が見つかった場合はそこへ確実に戻る。
                                        if ($null -ne $nextTabId) {
                                            # アクティブタブIDを親のIDへ更新する。
                                            $global:ActiveTabId = $nextTabId
                                            # 紐づくWebView2コントロールも親のものへ切り替える。
                                            $global:WebViewCtrl = $parentCtrl
                                            Write-DebugLog -Message "[WebView2] 情報: ゴーストを削除し、親タブ ($nextTabId) へ自動復帰しました。" -Level Info
                                        } else {
                                            # 親タブが見つからない場合のフォールバック： 万が一親が見つからない場合、残存しているタブへ強制移行させる。
                                            $remainingTabs = @($global:Tabs.Keys)
                                            if ($remainingTabs.Count -gt 0) {
                                                # 残存タブの先頭[0]を新しいアクティブタブとして強制的に設定する。
                                                $global:ActiveTabId = $remainingTabs[0]
                                                $global:WebViewCtrl = $global:Tabs[$global:ActiveTabId].WebView
                                                Write-DebugLog -Message "[WebView2] 警告: 親タブ不明のため、残存タブ ($($global:ActiveTabId)) へ強制移行しました。" -Level Warning
                                            } else {
                                                # 残存タブが存在しない場合、状態を完全にリセットする。
                                                $global:ActiveTabId = $null
                                                $global:WebViewCtrl = $null
                                                Write-DebugLog -Message "[WebView2] 情報: 全タブが閉じられたため、認識をリセットしました。" -Level Info
                                            }
                                        }
                                    }
                                } catch {
                                    Write-DebugLog -Message "[WebView2] 警告: ゴーストタブの破棄中にエラー ($($_.Exception.Message))" -Level Warning
                                }
                            }
                            # メインUIスレッドへ非同期処理(BeginInvoke)を依頼し、デリゲートを実行させる。
                            try { $global:BrowserForm.BeginInvoke($mi) | Out-Null } catch {}

                        <#
                          外部変数の状態（$tabIdや$parentCtrl）を別スレッド実行時でも正確に保持するため、
                          現在のスコープを完全にキャプチャする GetNewClosure() を必ず付与してイベントを登録する。
                        #>
                        }.GetNewClosure())

                        # サブ画面(ポップアップ)専用のクラッシュ検知： ポップアップタブのプロセス異常終了を検知する。
                        $senderInit.CoreWebView2.add_ProcessFailed({
                            param ($failSender, $failArgs)

                            # UIスレッド上で安全にコントロールを操作するため、MethodInvokerデリゲートを生成する。
                            $mi = [System.Windows.Forms.MethodInvoker]{
                                try {
                                    # エラーの種類(クラッシュ理由)を取得する。
                                    $kind = $failArgs.ProcessFailedKind
                                    Write-DebugLog -Message "[WebView2] エラー: ポップアップタブ($tabId)のクラッシュを検知 (Kind: $kind)" -Level Error
                                    
                                    # クラッシュしたゴーストタブを強制破棄し、メモリを解放する。
                                    $global:BrowserForm.Controls.Remove($senderInit)
                                    # WebView2コントロールのメモリやリソースを即座に破棄する。
                                    $senderInit.Dispose()
                                    if ($global:Tabs.ContainsKey($tabId)) {
                                        # タブ管理用のハッシュテーブルから該当のキーと値を消去する。
                                        $global:Tabs.Remove($tabId)
                                    }
                                    
                                    # 親タブへの強制復帰： クラッシュしたタブが現在操作中のタブであったか判定する。
                                    if ($global:ActiveTabId -eq $tabId) {
                                        # 次の移行先となるタブIDの格納用変数を初期化する。
                                        $nextTabId = $null
                                        # 親コントロールと一致するものを探索する。
                                        foreach ($key in $global:Tabs.Keys) {
                                            if ($global:Tabs[$key].WebView -eq $parentCtrl) { $nextTabId = $key; break }
                                        }
                                        # 親が見つかった場合、状態を親タブへ更新する。
                                        if ($null -ne $nextTabId) {
                                            $global:ActiveTabId = $nextTabId
                                            $global:WebViewCtrl = $parentCtrl
                                        }
                                    }
                                } catch {}
                            }
                            # メインUIスレッドへ非同期処理(BeginInvoke)を依頼し、デリゲートを実行させる。
                            try { $global:BrowserForm.BeginInvoke($mi) | Out-Null } catch {}

                        <#
                          外部変数の状態（$tabIdや$parentCtrl）を別スレッド実行時でも正確に保持するため、
                          現在のスコープを完全にキャプチャする GetNewClosure() を必ず付与してイベントを登録する。
                        #>
                        }.GetNewClosure())
                    } else {
                        # フェイルセーフ： サブ画面の初期化失敗時の安全処理を行う。
                        $exMsg = "理由不明"
                        try {
                            if ($null -ne $argsInit.InitializationException) {
                                $exMsg = $argsInit.InitializationException.Message
                            }
                        } catch {}
                        Write-DebugLog -Message "[WebView2] エラー: サブ画面の初期化に失敗 ($exMsg)" -Level Error
                        
                        # 初期化に失敗したWebView2コントロールを破棄し、メモリリーク(ゴースト化)を防止する。
                        try { 
                            if ($null -ne $senderInit) {
                                $senderInit.Dispose() 
                            }
                        } catch {}

                        # OS標準ブラウザでの予期せぬポップアップ起動を完全にブロックしつつ、処理保留を解除する。
                        try {
                            $evtArgs.Handled = $true
                            $deferral.Complete()
                        } catch {}
                    }
                } catch {
                    Write-DebugLog -Message "[WebView2] エラー: 初期化完了イベント内で例外 ($($_.Exception.Message))" -Level Error
                    try { $deferral.Complete() } catch {}
                }
            })
            # 新規作成したWebView2コントロール(仮想タブ)のブラウザコアを、親画面と同じ環境(セッション・Cookie等)で非同期に初期化する。
            $newWebView.EnsureCoreWebView2Async($sender.Environment) | Out-Null
        } catch {
            Write-DebugLog -Message "[WebView2] エラー: 新規ウィンドウ捕獲失敗 ($($_.Exception.Message))" -Level Error
        }
    }
    # メインのブラウザ画面に対し、上で定義した「ポップアップ横取り(NewWindowRequested)」のイベントハンドラを正式に登録（紐付け）する。
    $global:WebViewCtrl.CoreWebView2.add_NewWindowRequested($global:NewWindowHandler)

} catch {
    Write-DebugLog -Message "[System] 致命的エラー: WebView2 初期化失敗 ($($_.Exception.Message))" -Level Fatal
    exit 1
}

# [ウィンドウの最前面表示およびフォーカス確保]： Win32APIを利用し、ウィンドウをフォアグラウンドへ引き上げる。
try {
    # フォームのOSレベルのウィンドウハンドルを取得する。
    $hWnd = $global:BrowserForm.Handle
    # Win32APIを利用し、ウィンドウを元のサイズ(9=SW_RESTORE)で表示させる。
    [Win32Api.Win32Utils]::ShowWindow($hWnd, 9) | Out-Null
    # Win32APIを利用し、ウィンドウをフォアグラウンド(最前面)に引き上げる。
    [Win32Api.Win32Utils]::SetForegroundWindow($hWnd) | Out-Null

    # .NETのフォームActivateメソッドを呼び出し、アクティブ状態にする。
    $global:BrowserForm.Activate()
    # WebView2コントロール自体へフォーカスを設定し、キー入力を受け付ける状態にする。
    [void]$global:WebViewCtrl.Focus()

    # TopMost トグルを利用したフォーカスハック： Windowsの仕様上、他のアプリがアクティブ権限を握っている場合でも、確実にブラウザウィンドウを手前に引き上げる。
    $global:BrowserForm.TopMost = $true
    Start-Sleep -Milliseconds 80
    $global:BrowserForm.TopMost = $false

    try {
        # JavaScriptを流し込み、DOM内部のwindowオブジェクトとアクティブ要素へフォーカスを当てる。
        $global:WebViewCtrl.CoreWebView2.ExecuteScriptAsync(
            "window.focus(); if(document.activeElement) { document.activeElement.focus(); }"
        ) | Out-Null
    } catch {}

    # 画面遷移(ナビゲーション)が完了するたびに、フォーカスを再確保するイベントを登録する。
    $global:WebViewCtrl.CoreWebView2.add_NavigationCompleted({
        param ($sender, $e)

        # UIスレッド上で安全にコントロールを操作するため、MethodInvokerデリゲートを生成する。
        $mi = [System.Windows.Forms.MethodInvoker]{
            try {
                # 再度、Win32APIを利用してウィンドウをフォアグラウンドに引き上げる。
                [Win32Api.Win32Utils]::SetForegroundWindow($global:BrowserForm.Handle) | Out-Null
                $global:BrowserForm.Activate()
                [void]$global:WebViewCtrl.Focus()

                try {
                    # JavaScriptを利用し、DOM内部のフォーカスを再度設定する。
                    $global:WebViewCtrl.CoreWebView2.ExecuteScriptAsync(
                        "window.focus(); if(document.activeElement) { document.activeElement.focus(); }"
                    ) | Out-Null
                } catch {}
            } catch {}
        }
        # メインUIスレッドへ非同期処理を依頼する。
        try { $global:BrowserForm.BeginInvoke($mi) | Out-Null } catch {}
    })

    # メインブラウザプロセス自体のクラッシュ(異常終了)を検知するイベントを登録する。
    $global:WebViewCtrl.CoreWebView2.add_ProcessFailed({
        param ($sender, $e)

        try {
            # クラッシュの種類と理由を取得する。
            $kind = $e.ProcessFailedKind
            $reason = ""
            try { $reason = $e.Reason } catch {}

            $errMsg = "ブラウザプロセスのクラッシュ検知 (Kind: $kind, Reason: $reason)"
            Write-DebugLog -Message "[System] 致命的エラー: $errMsg" -Level Fatal

            try { [Console]::WriteLine("[ERROR] $errMsg") } catch {}

            # UIスレッド上で安全にフォームを閉じるため、MethodInvokerデリゲートを生成する。
            $mi = [System.Windows.Forms.MethodInvoker]{
                # ブラウザフォームを完全に閉じる(Close)。
                try { $global:BrowserForm.Close() } catch {}
            }
            # メインUIスレッドへ非同期処理を依頼する。
            try { $global:BrowserForm.BeginInvoke($mi) | Out-Null } catch {}
        } catch {}
    })

} catch {
    Write-DebugLog -Message "[System] 警告: フォーカス確保処理で例外発生 ($($_.Exception.Message))" -Level Warning
}

# --- WebView2プロファイルのキャッシュ等の削除 ---
# UDFのキャッシュ、Cookie、LocalStorage等を非同期で完全削除する。
function Clear-WebCache {
    param ([string]$Mode = "CacheOnly")

    $func = $MyInvocation.MyCommand.Name

    try {
        # アクティブなWebView2インスタンスを取得する。
        $webview = Get-ActiveWebView
        if ($null -eq $webview -or $null -eq $webview.CoreWebView2) {
            throw (New-EngineException -Func $func -Type "初期化エラー" -Message "WebView2が初期化されていないか、アクティブタブが存在しません")
        }

        # キャッシュ、ディスクキャッシュ、メモリキャッシュのフラグをビット論理和(-bor)で結合し、削除対象とする。
        $kinds = [Microsoft.Web.WebView2.Core.CoreWebView2BrowsingDataKinds]::Cache -bor
                 [Microsoft.Web.WebView2.Core.CoreWebView2BrowsingDataKinds]::DiskCache -bor
                 [Microsoft.Web.WebView2.Core.CoreWebView2BrowsingDataKinds]::MemoryCaches

        # 削除モードが"All"の場合、Cookieやストレージ群も削除対象に追加する。
        if ($Mode -eq "All") {
            # 既存のフラグへ、Cookie、ローカルストレージ、DOMストレージのフラグをビット論理和(-bor)で追加する。
            $kinds = $kinds -bor
                     [Microsoft.Web.WebView2.Core.CoreWebView2BrowsingDataKinds]::Cookies -bor
                     [Microsoft.Web.WebView2.Core.CoreWebView2BrowsingDataKinds]::LocalStorage -bor
                     [Microsoft.Web.WebView2.Core.CoreWebView2BrowsingDataKinds]::AllDomStorage
        }

        # プロファイルに対して、指定したデータのクリア処理(ClearBrowsingDataAsync)を非同期で開始する。
        $task = $webview.CoreWebView2.Profile.ClearBrowsingDataAsync($kinds)

        # タイムアウト監視用にストップウォッチを起動する。
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        $timeoutSec = 15

        # クリアタスクが完了するまで待機ループを回す。
        while (-not $task.IsCompleted) {
            if ($sw.Elapsed.TotalSeconds -gt $timeoutSec) {
                throw (New-EngineException -Func $func -Type "Timeout" -Message "キャッシュクリア処理がタイムアウトしました" -Details "${timeoutSec}秒経過")
            }
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 20
        }
        return "SUCCESS"

    } catch {
        $errMsg = $_.Exception.Message -replace "`r`n", " " -replace "`n", " "
        throw (New-EngineException -Func $func -Type "内部エラー" -Message "キャッシュクリア処理中に予期せぬエラーが発生しました" -Details $errMsg)
    }
}
