# ------------------------------------------------------------------------------#●#✘ 
<#
  [UIAccessibility (UIA) API と WebView2 (DOM) を連携させた要素ピッカー]： UIA情報をもとに
  ブラウザ内から最適なCSSセレクタを算出・逆引きし、自動化コードを生成する。
#>
# ------------------------------------------------------------------------------

# --- DOM逆引きコア ---
# UIA情報をもとにブラウザ内から最適なCSSセレクタを算出する（位置ベース優先）。
function Get-WebSelectorFromUia {
    param(
        [string]$UiaName,
        [string]$UiaType,
        [string]$UiaId,
        [string]$UiaHelpText,
        [System.Windows.Automation.AutomationElement]$UiaElement
    )

    $func = $MyInvocation.MyCommand.Name

    # 1. UIA (OSレベル) の絶対座標計算： UIA要素のOS画面上での境界短形(BoundingRectangle)を取得する。
    # UIA要素のOS画面上での境界短形(BoundingRectangle)を取得する。
    $rect = $UiaElement.Current.BoundingRectangle
    # 要素の物理的な中心X座標、中心Y座標を算出する。
    $screenX = $rect.Left + ($rect.Width / 2.0)
    $screenY = $rect.Top  + ($rect.Height / 2.0)

    # 2. OS絶対座標からブラウザ相対座標(Client座標)への変換： WebView2コントロールを基準とした相対座標へ変換する。
    $clientX = $screenX
    $clientY = $screenY
    if ($null -ne $global:WebViewCtrl) {
        try {
            # System.Drawingアセンブリが未ロードであれば読み込む。
            if (-not ('System.Drawing.Point' -as [type])) { Add-Type -AssemblyName System.Drawing }
            # 算出したOS座標からPointオブジェクトを生成する。
            $pt = New-Object System.Drawing.Point([int]$screenX, [int]$screenY)
            # WebView2コントロールの左上を原点(0,0)とした相対座標に変換する。
            $clientPt = $global:WebViewCtrl.PointToClient($pt)
            $clientX = $clientPt.X
            $clientY = $clientPt.Y
        } catch {}
    }
    Write-DebugLog -Message "[UiaMatch] Screen(X=$screenX, Y=$screenY) -> Client(X=$clientX, Y=$clientY)" -Level Info

    # 3. DPR(画面拡大率)とスクロール量の取得およびUIA座標の補正： ブラウザ内部のDPRとスクロール量を取得し、DOM絶対座標を算出する。
    $dpr = 1.0
    $scrollX = 0
    $scrollY = 0
    try {
        # ブラウザ内部のDPRおよび現在のスクロール量を取得するJSを定義する。
        $jsMetrics = "return JSON.stringify({ dpr: window.devicePixelRatio || 1, scrollX: window.scrollX || 0, scrollY: window.scrollY || 0 });"
        # JSを実行し、結果をPowerShellオブジェクトへデコードする。
        $metricsRes = Invoke-WebScript -Js $jsMetrics
        if (-not [string]::IsNullOrWhiteSpace($metricsRes)) {
            $metrics = $metricsRes | ConvertFrom-Json
            if ($metrics.dpr -gt 0) { $dpr = $metrics.dpr }
            $scrollX = $metrics.scrollX
            $scrollY = $metrics.scrollY
        }
    } catch {}

    # DOMとの比較用座標の算出： 物理ピクセルを論理ピクセルへ戻し、スクロール量を足してDOM全体での絶対座標とする。
    $uiaCenterX = ($clientX / $dpr) + $scrollX
    $uiaCenterY = ($clientY / $dpr) + $scrollY
    Write-DebugLog -Message "[UiaMatch] Dpr=$dpr, Scroll(X=$scrollX, Y=$scrollY), Target(X=$uiaCenterX, Y=$uiaCenterY)" -Level Info

    # 4. 事前取得した DOMSnapshot の展開： ピッカー起動時に記録しておいた全DOM要素の情報を取得する。
    $domNodes = $global:DomNodes
    if (-not $domNodes -or $domNodes.Count -eq 0) {
        Write-DebugLog -Message "[$func] 警告: 比較用のDOMSnapshotが存在しません" -Level Warn
        return $null
    }

    # 5. 位置ベース UIA → DOM マッチング (最優先ルート)： 座標のユークリッド距離を用いて最も近い要素を特定する。
    $bestNode = $null
    $bestScore = [double]::NegativeInfinity

    foreach ($node in $domNodes) {
        # 座標情報を持たない非表示要素などはスキップする。
        if (-not $node.boundingBox) { continue }

        $bb = $node.boundingBox
        # DOM要素の中心座標を算出する。
        $domCenterX = $bb.x + ($bb.width / 2.0)
        $domCenterY = $bb.y + ($bb.height / 2.0)

        # ユークリッド距離による近接度計算： X軸とY軸の差分から、直線距離をピクセル単位で算出する。
        $dx = $domCenterX - $uiaCenterX
        $dy = $domCenterY - $uiaCenterY
        $distance = [math]::Sqrt($dx * $dx + $dy * $dy)

        # UIA座標とDOM座標が大きく乖離している(150px以上)場合は誤爆とみなしてスキップする。
        if ($distance -gt 150) { continue }
        # 距離が近いほどスコアが高くなるように反比例の式を適用する。
        $score = 1000.0 / (1.0 + $distance)
        
        # RPAの操作対象になりやすいタグを優遇加点する。
        if ($node.tagName -eq "input") { $score += 100 }
        if ($node.tagName -eq "button") { $score += 50 }
        
        # プレースホルダの一致判定： UIA側でHelpTextとして認識されるため、一致すれば同一要素の可能性が高い。
        if ($node.placeholder -and $UiaHelpText -and $node.placeholder -eq $UiaHelpText) {
            $score += 200
        }

        # スコアがこれまでの最高値を上回ったら、候補を更新する。
        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestNode = $node
        }
    }

    # 6. 位置マッチング成功時のセレクタ生成： スコアに基づく有力候補からCSSセレクタを生成する。
    if ($bestNode) {
        Write-DebugLog -Message "[UiaMatch] 有力候補: Tag=$($bestNode.tagName), Id=$($bestNode.id), Name=$($bestNode.name) (Score: $([math]::Round($bestScore, 2)))" -Level Info
        # Nodeに既に完全なセレクタ文字列が保持されていればそれを採用する。            
        if ($bestNode.selector) { return $bestNode.selector }
        
        # 堅牢なセレクタ(IDとNameの複合)を生成する。
        $tagName = if ($bestNode.tagName) { $bestNode.tagName } else { "*" }
        if ($bestNode.id -and $bestNode.name) {
            return "$tagName#$($bestNode.id)[name=`"$($bestNode.name)`"]"
        }
        if ($bestNode.name)     { return ($tagName + '[name="' + $bestNode.name + '"]') }
        if ($bestNode.id)       { return ('#' + $bestNode.id) }
    }

    # 7. フォールバック： レーベンシュタイン距離による類似テキスト検索へ移行する。
    Write-DebugLog -Message "[$func] 情報: 有効な候補が見つからず、Fuzzy検索へフォールバックします" -Level Info
    $cssSelector = Get-WebSelectorFromUiaFuzzy -UiaName $UiaName -UiaType $UiaType -UiaId $UiaId -UiaHelpText $UiaHelpText
    if ($cssSelector) { return $cssSelector }

    return $null
}

# ------------------------------------------------------------------------------
# --- フォールバック解析 --- 
# レーベンシュタイン距離を用いた曖昧 (Fuzzy) 検索を実行する。
function Get-WebSelectorFromUiaFuzzy {
    param(
        [string]$UiaName,
        [string]$UiaType,
        [string]$UiaId,
        [string]$UiaHelpText
    )

    $domNodes = $global:DomNodes
    if (-not $domNodes -or $domNodes.Count -eq 0) { return $null }

    # レーベンシュタイン距離の高速計算： 1次元配列を使い回してメモリ割り当てを抑制する。
    function Get-LevenshteinDistanceFast {
        param([string]$s, [string]$t)

        # 片方が空文字列の場合は、もう一方の長さをそのまま距離として返す。
        if ([string]::IsNullOrEmpty($s)) { return $t.Length }
        if ([string]::IsNullOrEmpty($t)) { return $s.Length }
        # 計算用の1次元配列(行のバッファ)を2つ用意する。
        $v0 = New-Object int[] ($t.Length + 1)
        $v1 = New-Object int[] ($t.Length + 1)

        # 初期化： 空文字列からtへの変換コスト(純粋な追加コスト)を計算する。
        for ($i = 0; $i -le $t.Length; $i++) { $v0[$i] = $i }
        # 動的計画法(DP)によるマトリックス計算を開始する。
        for ($i = 0; $i -lt $s.Length; $i++) {
            # 行の先頭は、空文字列からsへの変換コストとする。
            $v1[0] = $i + 1
            for ($j = 0; $j -lt $t.Length; $j++) {
                # 文字が一致していれば置換コストは0、異なれば1とする。
                $cost = if ($s[$i] -eq $t[$j]) { 0 } else { 1 }
                # 削除、挿入、置換の中で最もコストの小さい経路を選択する。
                $v1[$j + 1] = [Math]::Min([Math]::Min($v1[$j] + 1, $v0[$j + 1] + 1), ($v0[$j] + $cost))
            }
            # 次の行の計算に向けて配列状態をコピーする。
            for ($j = 0; $j -le $t.Length; $j++) { $v0[$j] = $v1[$j] }
        }
        return $v1[$t.Length]
    }

    $bestNode = $null
    $bestScore = [double]::NegativeInfinity
    
    # UIA側の情報から比較用のベース文字列を生成する。
    $uiaText = ("$UiaName $UiaHelpText $UiaId $UiaType").Trim()
    if ([string]::IsNullOrWhiteSpace($uiaText)) { return $null }
    
    # 探索対象を操作可能な要素に絞り込み、計算コストを削減する。
    $interactiveTags = @("input", "button", "a", "select", "textarea", "label", "span", "div", "img")

    foreach ($node in $domNodes) {
        if (-not $node.tagName -or $node.tagName -notin $interactiveTags) { continue }
        # DOM側の情報からも比較用のベース文字列を生成する。
        $domText = "$($node.placeholder) $($node.name) $($node.id) $($node.tagName) $($node.innerText)".Trim()
        if (-not $domText) { continue }

        # --- 新スコアリングロジック ---
        $score = 0

        # 1. IDまたはNameの直接比較： UIAのID・NameとDOM属性を比較しスコアを加算する。
        if ($UiaId -and $node.id) {
            # IDが完全一致する場合は超高得点を付与する。
            if ($node.id -eq $UiaId) { $score += 800 }
            # IDが部分一致(包含関係)の場合は中得点を付与する。
            elseif ($node.id -match $UiaId -or $UiaId -match $node.id) { $score += 300 }
        }
        if ($UiaName) {
            # Nameやプレースホルダが完全一致する場合は高得点を付与する。
            if ($node.name -eq $UiaName -or $node.placeholder -eq $UiaName) { $score += 600 }
            # innerText(画面表示テキスト)が一致した場合の大幅加点（歳出管理などを救済）
            elseif ($node.innerText -and $node.innerText -eq $UiaName) { $score += 600 }
            elseif ($node.innerText -and $node.innerText -match $UiaName) { $score += 200 }
            # 部分一致の場合は加点する。
            elseif (($node.name -and $node.name -match $UiaName) -or ($node.placeholder -and $node.placeholder -match $UiaName)) { $score += 200 }
            elseif (($node.name -and $UiaName -match $node.name) -or ($node.placeholder -and $UiaName -match $node.placeholder)) { $score += 200 }
        }

        # 2. レーベンシュタイン距離を補助スコアとして加算： 類似文字列の編集距離から加点スコアを算出する。
        # 編集距離が小さい（文字列が似ている）ほど加点が大きくなるように反比例の式を適用する。
        $dist = Get-LevenshteinDistanceFast -s $uiaText -t $domText
        $score += (500.0 / (1.0 + $dist))

        # 3. RPAで操作されやすいタグの優遇： 入力要素やボタンに加点する。
        if ($node.tagName -eq "input") { $score += 50 }
        if ($node.tagName -eq "button") { $score += 40 }
        # スコアがこれまでの最高値を上回ったら、候補を更新する。
        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestNode = $node
        }
    }

    # スコアの足切りライン(100点)を設け、無関係なinputタグの誤爆を防ぐ
    if ($bestNode -and $bestScore -ge 100) {
        if ($bestNode.selector) { return $bestNode.selector }
        if ($bestNode.name)     { return $bestNode.tagName + '[name="' + $bestNode.name + '"]' }
        if ($bestNode.id)       { return '#' + $bestNode.id }
    }
    return $null
}

# ==============================================================================
# --- 要素ピッカーの実行要求窓口 ---
# マウスポインタ下の要素を特定し、最適なCSSセレクタとRPAコマンドを自動生成する（VBAエントリーポイント）。
# ==============================================================================
function Invoke-UiaRecordStep {
    param (
        [int]$WaitSec = 3,
        [bool]$UsePointReverse = $false
    )

    $func = $MyInvocation.MyCommand.Name
    Write-DebugLog -Message "================================================================================" -Level Info
<#
  座標逆引き（$UsePointReverse = $false)、VBAフォーム側で &true (規定値)
  Fuzzy検索は、要素が innerText や aria-label などの「何かしらの文字列」を持っていることが前提、しかし、
  モダンなWebアプリでは「虫眼鏡アイコンだけの検索ボタン（SVG画像のみ）」やハンバーガーメニューなど、テキストを持たない要素が多々ある。
  UIA側にもDOM側にもテキストが存在しない場合、Fuzzy検索は空振りする。この時、マウスの物理的な位置から要素を取得する。
  • 基本はスイッチOFF（UIA-DOM / Fuzzyルート）,,, 「きれいなCSSセレクタ」が生成されやすい。
  • スイッチON（PointReverseルート）   ,,,   XPATH://a[contains(...)] 絶対パスなど「確実なセレクタ」
#>

    try {
        # WinForms(マウスポインタ座標取得用)のロード
        if (-not ('System.Windows.Forms.Cursor' -as [type])) {
            Add-Type -AssemblyName System.Windows.Forms
        }
        
        # ユーザーが対象要素にマウスを合わせるための待機: （現在 $WaitSec=0、VBA側で制御している。）
        if ($WaitSec -gt 0) {
            Start-Sleep -Seconds $WaitSec
        }
        
        # 待機完了後に、スクロール後を捉えるため最新のDOMスナップショットを取得する。
        $global:DomNodes = Get-DomSnapshotWithBoxModel

        # メイン処理： 要素特定オーケストレーターへ処理を委譲する。
        $generatedCode = Process-RecordAction -IsGetNameMode $false -UsePointReverse $UsePointReverse

        if ([string]::IsNullOrWhiteSpace($generatedCode)) {
            throw (New-EngineException -Func $func -Type "未発見" -Message "対象要素の取得、またはコードの生成に失敗しました")
        }
        return $generatedCode

    } catch {
        # VBA側に On Error Resume Next がある前提で例外を投げる。
        throw (New-EngineException -Func $func -Type "UIAエラー" -Message "UIA要素の記録処理中にエラーが発生しました" -Details $_.Exception.Message)
    }
}

# ------------------------------------------------------------------------------
# --- 共通JSコア --- 
# 物理座標から要素を逆引きするShadow DOM貫通ロジックをJSとして定義する。
function Get-CorePointReverseJs {
    param(
        [int]$ClientX,
        [int]$ClientY
    )

    $clientXLocal = $ClientX
    $clientYLocal = $ClientY
    return @"
    try {
        // PowerShellから渡されたブラウザ相対座標を変数へ格納する。
        let targetX = $clientXLocal;
        let targetY = $clientYLocal;
        // OSのディスプレイ拡大率(DPR)を取得する。
        const dpr = window.devicePixelRatio || 1;

        // 物理ピクセル座標を、ブラウザ内部の論理ピクセル座標(CSSピクセル)へ変換する。
        targetX = targetX / dpr;
        targetY = targetY / dpr;

        // 座標から最も深い要素(Shadow DOM / iframeを貫通)を特定する再帰関数
        function deepElementFromPoint(x, y, rootNode) {
            // Shadow DOM内部を探索するための補助関数
            function probeShadow(root, px, py) {
                try {
                    if (!root) return null;
                    // ルート内で指定座標にある要素を取得する。
                    if (root.elementFromPoint) {
                        const el = root.elementFromPoint(px, py);
                        if (el) return el;
                    }
                    // APIが利用できない場合のフォールバックとして、操作可能な主要要素を抽出する。
                    const candidates = root.querySelectorAll ? root.querySelectorAll('input,button,textarea,select,a,[role="search"],[role="textbox"]') : [];
                    if (candidates && candidates.length) return candidates[0];
                } catch (e) {}
                return null;
            }

            let stacked = [];
            try {
                // 指定座標に重なっているすべての要素を配列(stacked)として取得する。ブラウザの互換性を考慮して複数APIを試行する。
                if (rootNode.elementsFromPoint) {
                    stacked = rootNode.elementsFromPoint(x, y);
                } else if (document.elementsFromPoint) {
                    stacked = document.elementsFromPoint(x, y);
                } else if (document.msElementsFromPoint) {
                    stacked = document.msElementsFromPoint(x, y);
                } else {
                    // 重なり要素が取得できない場合は、最前面の単一要素のみを配列に格納する。
                    const el = (rootNode.elementFromPoint || document.elementFromPoint).call(rootNode || document, x, y);
                    if (el) stacked = [el];
                }
            } catch (e) {
                stacked = [];
            }

            // 取得した重なり要素を前面(Z-Indexが高い)順に検証する。
            for (let i = 0; i < stacked.length; i++) {
                const el = stacked[i];
                if (!el) continue;

                // 隠しフィールド(type="hidden")は操作不能なためスキップする。
                if (el.tagName === 'INPUT' && el.type === 'hidden') continue;
                // --- RPAの主要操作対象の特定 (優先順位付き) ---
             
                // 優先度1: contenteditableな要素 (TinyMCE等のリッチテキストエディタのBodyを最優先)
                if (el.isContentEditable) return el;
             
                // 優先度2: 一般的なフォーム入力要素
                if (/^(INPUT|BUTTON|TEXTAREA|SELECT)$/.test(el.tagName)) return el;
             
                // 優先度3: Aタグ (透明な巨大オーバーレイの誤爆を防ぐため、テキストかhrefが実在するものに限定)
                if (el.tagName === 'A') {
                    if (!el.innerText.trim() && !el.getAttribute('aria-label') && !el.href) continue;
                    if (el.classList.contains('anchor-link')) continue; 
                    return el;
                }
             
                // 優先度4: ARIAロールによるモダンWeb(SPA)のカスタムボタン・テキストボックス等
                const elRole = el.getAttribute ? el.getAttribute('role') : null;
                if (elRole && /^(button|link|checkbox|menuitem|tab|searchbox|textbox)$/i.test(elRole)) return el;
                // LABELタグがクリックされた場合、関連付けられた実際の入力フォーム(Input等)へ対象をすり替える。
                if (el.tagName === 'LABEL') {
                    const forId = el.getAttribute('for');
                    if (forId) {
                        const linked = document.getElementById(forId);
                        if (linked) return linked;
                    }
                    const innerInput = el.querySelector('input,textarea,select');
                    if (innerInput) return innerInput;
                }

                // マウス位置の要素がデザイン用の DIV や SPAN だった場合、その親要素の枠内に、本物の SELECT や INPUT が隠れていないか探索してすり替える。
                if (/^(DIV|SPAN|I|SVG|P|B)$/.test(el.tagName) && el.parentElement) {
                    try {
                        const siblingTarget = el.parentElement.querySelector('select, input:not([type="hidden"]), textarea, button');
                        // 見つかった場合、それが確実にRPAで操作したい本体要素なので優先する
                        if (siblingTarget) return siblingTarget;
                    } catch(e) {}
                }

                // 要素の内部に操作対象が含まれている場合(DIVで包まれたボタン等)の抽出処理
                try {
                    const inner = el.querySelector && el.querySelector('input,textarea,select,button,a,[role="textbox"],[role="search"]');
                    if (inner) return inner;
                } catch (e) {}

                // Shadow DOMのさらに内部にあるスロット(SLOT)要素の透過処理
                try {
                    if (el.tagName === 'SLOT') {
                        const assigned = el.assignedElements ? el.assignedElements() : [];
                        if (assigned && assigned.length) {
                            for (const a of assigned) {
                                const found = deepElementFromPoint(x, y, a.ownerDocument || document);
                                if (found) return found;
                            }
                        }
                    }
                } catch (e) {}

                // 要素がShadow DOM(カプセル化されたコンポーネント)を持っている場合の透過処理
                try {
                    if (el.shadowRoot) {
                        const sEl = probeShadow(el.shadowRoot, x, y);
                        if (sEl) {
                            // 多段Shadow DOMに対応するため再帰呼び出しを行う。
                            if (sEl.shadowRoot) {
                                const deep = deepElementFromPoint(x, y, sEl.shadowRoot);
                                if (deep) return deep;
                            }
                            return sEl;
                        }

                        // 別の抽出手段へのフォールバック
                        try {
                            const inner = el.shadowRoot.elementFromPoint ? el.shadowRoot.elementFromPoint(x, y) : null;
                            if (inner && inner !== el) {
                                if (inner.shadowRoot) {
                                    const deep = deepElementFromPoint(x, y, inner.shadowRoot);
                                    if (deep) return deep;
                                }
                                const innerInput = inner.querySelector && inner.querySelector('input,textarea,select,button,a,[role="textbox"],[role="search"]');
                                if (innerInput) return innerInput;
                                return inner;
                            }
                        } catch (e) {}
                    }
                } catch (e) {}

                // iframeまたはframe内部の要素である場合、座標を内部相対値に再計算し、透過して探索を継続する。
                try {
                    if (el.tagName === 'IFRAME' || el.tagName === 'FRAME') {
                        const rect = el.getBoundingClientRect();
                        const iframeX = x - rect.left;
                        const iframeY = y - rect.top;
                        const doc = el.contentDocument;
                        if (doc) {
                            const inner = deepElementFromPoint(iframeX, iframeY, doc);
                            if (inner) return inner;
                        }
                    }
                } catch (e) {}
            }

            // 上記の判定ですべて弾かれた場合、通常の最前面要素を強制的に採用する(最終フォールバック)
            try {
                const fallback = (rootNode.elementFromPoint || document.elementFromPoint).call(rootNode || document, x, y);
                if (fallback) {
                    if (fallback.tagName === 'LABEL') {
                        const forId = fallback.getAttribute && fallback.getAttribute('for');
                        if (forId) {
                            const linked = document.getElementById(forId);
                            if (linked) return linked;
                        }
                        const innerInput = fallback.querySelector && fallback.querySelector('input,textarea,select');
                        if (innerInput) return innerInput;
                    }
                    return fallback;
                }
            } catch (e) {}

            return null;
        }

        // 特定した要素から一意なCSSセレクタを構築する関数
        function buildCssSelector(el) {
            if (!el) return '';
            // CSS.escapeが存在すれば利用し、無ければ何もしない無名関数を定義する。
            const cssEsc = (typeof CSS !== 'undefined' && CSS.escape) ? CSS.escape : s => s;

            // IDが存在する場合は最優先でIDセレクタを採用する。
            if (el.id) {
                const idSel = '#' + cssEsc(el.id);
                try { if (document.querySelectorAll(idSel).length === 1) return idSel; } catch(e) {}
                // ID単体で一意にならなければ、name属性と組み合わせてみる。
                if (el.name) return el.tagName.toLowerCase() + idSel + '[name="' + cssEsc(el.name) + '"]';
            }
            // name属性が存在する場合は採用する。
            if (el.name) return el.tagName.toLowerCase() + '[name="' + cssEsc(el.name) + '"]';
            // クラスが存在する場合は結合してクラスセレクタを構築する。
            if (el.className && typeof el.className === 'string') {
                const classes = el.className.trim().split(/\s+/).map(cls => cssEsc(cls)).join('.');
                if (classes) return el.tagName.toLowerCase() + '.' + classes;
            }
            return '';
        }

        // CSSセレクタで一意に特定できない場合のための堅牢なXPathを構築する関数
        function buildRobustXPath(el) {
            if (!el) return '';
            // ID、Name属性での単独指定（XPath版）
            if (el.id) return '//' + el.tagName.toLowerCase() + '[@id="' + el.id + '"]';
            if (el.name) return '//' + el.tagName.toLowerCase() + '[@name="' + el.name + '"]';

            // ボタンやリンクの場合、内部のテキストを利用した結合検索(contains)を試みる。
            const text = (el.innerText || el.textContent || '').trim();
            if (text && text.length > 0 && text.length < 30 && (el.tagName === 'BUTTON' || el.tagName === 'A')) {
                if (text.indexOf("'") === -1 && text.indexOf('"') === -1) {
                    return '//' + el.tagName.toLowerCase() + '[contains(normalize-space(.), \'' + text + '\')]';
                }
            }

            // 上記のすべてでダメだった場合、ルートから階層を遡りながら絶対パスXPathを構築する。
            let xpath = '';
            let node = el;
            while (node && node.nodeType === 1) {
                if (node.id) {
                    // 親階層でIDを持つ要素が見つかれば、そこを起点にしてパスを短縮する。
                    xpath = '//' + node.tagName.toLowerCase() + '[@id="' + node.id + '"]' + xpath;
                    return xpath; 
                }
 
                 // 兄弟要素の中でのインデックス位置(1始まり)を計算する。               
                const siblings = node.parentNode ? node.parentNode.children : [];
                let count = 0, index = 1;
                for (let i = 0; i < siblings.length; i++) {
                    if (siblings[i].tagName === node.tagName) count++;
                    if (siblings[i] === node) { index = count; break; }
                }

                // インデックスを付加してパスを連結する。               
                const idxStr = (count > 1) ? '[' + index + ']' : '';
                xpath = '/' + node.tagName.toLowerCase() + idxStr + xpath;
                node = node.parentNode;
            }
            return xpath;
        }

        // メイン処理： 座標から要素を抽出する。
        const targetElement = deepElementFromPoint(targetX, targetY, document);

        // iframe内要素かどうかの確実な判定
        let isFrame = false;
        try {
            // targetElementが所属するドキュメントが、現在のトップドキュメント(document)と異なるかで判定
            isFrame = (targetElement && targetElement.ownerDocument !== document);
        } catch(e) {}

        // ピッカー操作の視覚的フィードバック
        if (targetElement && targetElement.style !== undefined) {
            try {
                const oldOutline = targetElement.style.outline;
                const oldOutlineOffset = targetElement.style.outlineOffset;
                const oldBackground = targetElement.style.backgroundColor; // 元の背景色を保存
                
                targetElement.style.outline = '2px solid red';
                targetElement.style.outlineOffset = '-2px'; 
                
                // iframe内の要素だった場合、背景色を薄黄色にして視覚的に警告する
                if (isFrame) {
                    targetElement.style.backgroundColor = '#ffffcc';
                }

                // 2秒後に枠線と背景色を元の状態へ復元する。
                setTimeout(function() { 
                    targetElement.style.outline = oldOutline; 
                    targetElement.style.outlineOffset = oldOutlineOffset;
                    if (isFrame) {
                        targetElement.style.backgroundColor = oldBackground;
                    }
                }, 2000); 
            } catch(e) {}
        }

        // デバッグ・追跡用のログオブジェクトを構築する。
        const log = {
            dpr: dpr,
            targetX: targetX,
            targetY: targetY,
            elementTag: targetElement ? targetElement.tagName : null,
            elementId: targetElement ? targetElement.id : null,
            elementName: targetElement ? targetElement.name : null
        };

        if (!targetElement) {
            return JSON.stringify({ error: 'elementFromPoint=null', log: log });
        }

        // セレクタの構築および判定。強固なCSSセレクタが得られなかった場合はXPATHへフォールバックする。
        const cssResult = buildCssSelector(targetElement);
        const finalSelector = (cssResult && (cssResult.indexOf('#') !== -1 || cssResult.indexOf('[name=') !== -1 || cssResult.indexOf('.') !== -1)) 
            ? cssResult : 'XPATH:' + buildRobustXPath(targetElement);

        // 要素の絶対座標（スクロール量加算済）を算出する。
        const rect = targetElement.getBoundingClientRect();
        const scrollX = window.scrollX || window.pageXOffset || 0;
        const scrollY = window.scrollY || window.pageYOffset || 0;

        // 最終的な解析結果をJSONとしてPowerShellへ返却する。
        return JSON.stringify({
            selector: finalSelector,
            outerHtml: targetElement.outerHTML || '',
            boundingBox: { 
                x: Math.round(rect.left + scrollX), 
                y: Math.round(rect.top + scrollY), 
                width: Math.round(rect.width), 
                height: Math.round(rect.height) 
            },
            isInsideFrame: isFrame, // 修正した判定結果を渡す
            log: log
        });
    } catch (e) {
        return JSON.stringify({ error: e && e.message ? e.message : String(e) });
    }
"@
}

# ------------------------------------------------------------------------------
# --- 物理座標ベースの要素逆引きおよび座標補正 ---
# ルートA： OS物理座標をブラウザ論理座標へ変換し、DOM要素を逆引き・補正する。
function Invoke-BrowserPointReverse {
    param(
        [System.Windows.Point]$ScreenPoint,
        [string]$TargetControlType,
        [string]$CurrentValue
    )

    # 1. スクリーン座標(OS)からクライアント座標(ブラウザ内)への変換： OS物理座標を相対座標へ変換する。
    $clientX = $ScreenPoint.X
    $clientY = $ScreenPoint.Y
    if ($null -ne $global:WebViewCtrl) {
        try {
            # System.Drawingが未ロードであれば読み込み、座標変換処理を実行する。
            if (-not ('System.Drawing.Point' -as [type])) { Add-Type -AssemblyName System.Drawing }
            $pt = New-Object System.Drawing.Point([int]$ScreenPoint.X, [int]$ScreenPoint.Y)
            $clientPt = $global:WebViewCtrl.PointToClient($pt)
            $clientX = $clientPt.X
            $clientY = $clientPt.Y
        } catch {}
    }
    Write-DebugLog -Message "[PointReverse] Screen(X=$($ScreenPoint.X), Y=$($ScreenPoint.Y)) -> Client(X=$clientX, Y=$clientY)" -Level Info

    # 2. ブラウザ内でのヒットテスト（共通JSロジック）： コアJSロジックを利用して要素を逆引きする。
    # コアJSロジック(Get-CorePointReverseJs)に変換済みのX/Y座標を埋め込んでJSコードを生成する。
    $jsPoint = Get-CorePointReverseJs -ClientX $clientX -ClientY $clientY
    # 生成したJSをブラウザで実行し、JSON結果を取得する。
    $resRaw = Invoke-WebScript -Js $jsPoint
    if (-not $resRaw) { return $null }
    # JSON文字列をPowerShellオブジェクトへ変換する。
    $domData = if ($resRaw -is [string]) { try { $resRaw | ConvertFrom-Json } catch { $null } } else { $resRaw }
    
    if ($domData.log) {
        Write-DebugLog -Message "[PointReverse] Dpr=$($domData.log.dpr), Target(X=$($domData.log.targetX), Y=$($domData.log.targetY))" -Level Info
        Write-DebugLog -Message "[PointReverse] elementFromPoint => Tag=$($domData.log.elementTag), Id=$($domData.log.elementId), Name=$($domData.log.elementName)" -Level Info
    }

    if ($null -eq $domData -or $domData.error) { return $null }

    # 3. 座標補正ロジック： JSで取得した座標はズレが生じやすいため、CDP(DevTools)で事前取得した正確なDOMスナップショットの座標で上書き補正する。
    if ($domData.boundingBox -and $global:DomNodes) {
        # セレクタが完全一致するノードをスナップショットから探す。
        $matchedNode = $global:DomNodes | Where-Object { $_.selector -and $_.selector -eq $domData.selector } | Select-Object -First 1

        if (-not $matchedNode -and $domData.log) {
            # セレクタで一致しない場合、タグとID・Nameの組み合わせで厳密に同定を試みる。
            $matchedNode = $global:DomNodes | Where-Object {
                $isMatch = $true
                $hasId = -not [string]::IsNullOrEmpty($domData.log.elementId)
                $hasName = -not [string]::IsNullOrEmpty($domData.log.elementName)
            
                # IDもNameも無い場合は確証が得られないため除外（誤爆防止）
                if (-not $hasId -and -not $hasName) { return $false }
            
                # 双方が存在する場合は両方の一致を必須とし(Yahoo路線等の重複ID対策)、
                # 片方のみの場合はその片方の一致を必須とする。
                if ($hasId -and $_.id -ne $domData.log.elementId) { $isMatch = $false }
                if ($hasName -and $_.name -ne $domData.log.elementName) { $isMatch = $false }
           
                return $isMatch
            } | Select-Object -First 1
        }

        # CDP座標による上書き： JSで取得した不完全なBoundingBoxを、CDP側の精密なBoundingBoxで上書きする。
        if ($matchedNode -and $matchedNode.boundingBox) {
            Write-DebugLog -Message "[PointReverse] 情報: DOMスナップショット(CDP)の正確なBoundingBoxで座標を補正(同期)します" -Level Info
            # JSで取得した不完全なBoundingBoxを、CDP側の精密なBoundingBoxで上書きする。
            $domData.boundingBox = $matchedNode.boundingBox
        }
    }
    Write-DebugLog -Message "[PointReverse] 確定結果: Selector='$($domData.selector)' | 座標(x:$($domData.boundingBox.x), y:$($domData.boundingBox.y))" -Level Info

    # 取得した情報から最終的なPowerShellコード文字列(RPAコマンド)を組み上げて返却する。
    $isFrame = if ($domData.isInsideFrame) { $true } else { $false }
    return (ConvertTo-RpaCommandCode -Selector $domData.selector -ControlType $TargetControlType -CurrentValue $CurrentValue -OuterHtml $domData.outerHtml -BoundingBox $domData.boundingBox -IsInsideFrame $isFrame)
}

# ------------------------------------------------------------------------------
# --- UIAメタデータからのDOM逆引き ---
# ルートB： UIAのテキスト情報をもとにDOM要素を特定する。
function Invoke-UiaDomResolution {
    param(
        [string]$TargetName,
        [string]$TargetControlType,
        [string]$TargetId,
        [System.Windows.Automation.AutomationElement]$TargetElement,
        [string]$CurrentValue
    )

    # UIA情報をもとに最適なCSSセレクタを算出する。
    $cssSelector = Get-WebSelectorFromUia -UiaName $TargetName -UiaType $TargetControlType -UiaId $TargetId -UiaHelpText "" -UiaElement $TargetElement
    if ([string]::IsNullOrWhiteSpace($cssSelector)) { return $null }

    # Base64エンコードによるサニタイズ
    $selB64 = ConvertTo-JsSafeBase64 -Text $cssSelector

    $scriptDom = @"
        try {
            // Base64をデコードしてセレクタ文字列を復元する。
            const sel = decodeURIComponent(escape(atob('$selB64')));
            // XPATH: から始まる場合はXPath評価、それ以外はquerySelectorを利用して要素を取得する。
            const el = sel.indexOf('XPATH:') === 0 
                ? document.evaluate(sel.substring(6), document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue 
                : document.querySelector(sel);
            if(!el) return null;

            // 座標補正： 隠されたInput要素の代わりにLabelの可視座標を抽出する。
            // 取得した要素のBoundingBoxとスタイル情報を取得する。
            let targetForBox = el;
            const rectCheck = el.getBoundingClientRect();
            const style = window.getComputedStyle(el);

            // チェックボックスやラジオボタンなど、実体(Input)が透明化・サイズ0にされて隠されている場合を検知する。            
            if (el.tagName === 'INPUT' && (rectCheck.width === 0 || rectCheck.height === 0 || style.opacity === '0' || style.position === 'absolute')) {
                try {
                    // IDが存在する場合、for属性で紐づいているlabel要素を探す。
                    if (el.id) {
                        const label = document.querySelector('label[for="' + el.id + '"]');
                        if (label) targetForBox = label;
                    }
                    // labelが見つからず、かつ親要素がlabelである場合(Inputを内包するLabel)は、親を座標取得の対象とする。
                    if (targetForBox === el && el.parentElement && el.parentElement.tagName === 'LABEL') {
                        targetForBox = el.parentElement;
                    }
                } catch(e) {}
            }

            // 補正された要素から座標を取得し、スクロール量を加算して絶対座標とする。
            const rect = targetForBox.getBoundingClientRect();
            const sx = window.scrollX || window.pageXOffset || 0;
            const sy = window.scrollY || window.pageYOffset || 0;
            // 外側のHTML構造と座標情報をJSON文字列で返却する。
            return JSON.stringify({
                outerHtml: el.outerHTML || '',
                boundingBox: { 
                    x: Math.round(rect.left + sx), 
                    y: Math.round(rect.top + sy), 
                    width: Math.round(rect.width), 
                    height: Math.round(rect.height) 
                }
            });
        } catch(e) { return null; }
"@

    # 構築したJSを実行し、要素の詳細情報を取得する。
    $resultDom = Invoke-WebScript -Js $scriptDom
    $parsedInfo = if ($resultDom -is [string]) { try { $resultDom | ConvertFrom-Json } catch { $null } } else { $resultDom }

    # 取得情報からHTML文字列とBoundingBoxを取り出す。
    $outHtml = if ($parsedInfo -and $parsedInfo.outerHtml) {$parsedInfo.outerHtml } else { "" }
    $bbox = if ($parsedInfo -and $parsedInfo.boundingBox) { $parsedInfo.boundingBox } else { $null }

    # 取得した情報をもとにRPAコマンド文字列を組み上げて返却する。
    return (ConvertTo-RpaCommandCode -Selector $cssSelector -ControlType $TargetControlType -CurrentValue $CurrentValue -OuterHtml $outHtml -BoundingBox $bbox)
}

# ------------------------------------------------------------------------------
# --- デスクトップ要素(UIAモード)へのフォールバック --- 
# ルートC： ブラウザ以外のネイティブWindowsアプリがターゲットになった場合のコードを生成する。
function Build-FallbackStepCode {
    param(
        [string]$WindowName,
        [string]$TargetName,
        [string]$ControlType,
        [string]$CurrentValue,
        [bool]$IsGetNameMode
    )

    # GetName(値の取得)モードの場合、TargetNameをワイルドカード(*)にして柔軟性を持たせる。
    $stepTargetName = if ($IsGetNameMode) { "*" } else { $TargetName }

    # ネイティブUIA用のコマンド(Execute-UIRpaStep)を組み立てる関数を呼び出す。
    $stepCode = Build-RpaStepCode `
        -Action "Click" `
        -WindowName $WindowName `
        -TargetName $stepTargetName `
        -ControlType $ControlType `
        -Index 1 `
        -ActionValue $CurrentValue `
        -AnchorName "" `
        -AnchorDirection "None" `
        -AnchorSteps 0

    # 文字列ではなく、ハッシュテーブルとして返却する。
    return @{
        command     = $stepCode
        selector    = ""
        value       = $CurrentValue
        outerHtml   = ""
        boundingBox = $null
        isInsideFrame = $false
    }
}

# ------------------------------------------------------------------------------
# --- 最終的なRPAコマンドの組み上げ --- 
# ヘルパー： 特定された情報をもとに、VBAから実行可能なRPAコマンド文字列(JSON)を最終出力する。
function ConvertTo-RpaCommandCode {
    param(
        [string]$Selector,
        [string]$ControlType,
        [string]$CurrentValue,
        [string]$OuterHtml,
        [object]$BoundingBox,
        [bool]$IsInsideFrame = $false
    )

    # 文字列内に含まれるダブルクォートをシングルクォートに置換し、改行をスペースにしてVBA側への受け渡しを安全にする。
    $outHtmlSafe = if ($OuterHtml) { $OuterHtml -replace '"', "'" -replace "`r`n|`n|`r", " " } else { "" }
    # 座標オブジェクト(BoundingBox)を圧縮されたJSON文字列に変換する。
    $bboxStr = if ($BoundingBox) { $BoundingBox | ConvertTo-Json -Compress -Depth 5 } else { "" }
    # 入力値に含まれるダブルクォートをPowerShellのエスケープシーケンス(`")に置換する。
    $safeValue = if ($CurrentValue) { $CurrentValue -replace '"', '`"' } else { "" }

    # コマンド文字列の生成： セレクタがXPath形式かどうかで発行するコマンド(WebXPath系かWeb系か)を分岐させる。
    $cmdStr = ""
    if ($Selector.StartsWith("XPATH:")) {
        $xp = $Selector.Substring(6)
        if ($ControlType -eq "Edit") {
            # テキスト入力ボックス(Edit)の場合、入力コマンド(Set-WebXPathTextInput)を発行する。
            $cmdStr = "Set-WebXPathTextInput -XPath `"$xp`" -Value `"$safeValue`" -OuterHtml `"$outHtmlSafe`" -BoundingBox `"$bboxStr`""
        } else {
            # それ以外はクリックコマンド(Invoke-WebXPathClick)を発行する。
            $cmdStr = "Invoke-WebXPathClick -XPath `"$xp`" -OuterHtml `"$outHtmlSafe`" -BoundingBox `"$bboxStr`""
        }
    } else {
        if ($ControlType -eq "Edit") {
            # CSSセレクタでのテキスト入力コマンド(Set-WebTextInput)を発行する。
            $cmdStr = "Set-WebTextInput -Selector `"$Selector`" -Value `"$safeValue`" -OuterHtml `"$outHtmlSafe`" -BoundingBox `"$bboxStr`""
        } else {
            # CSSセレクタでのクリックコマンド(Invoke-WebClick)を発行する。
            $cmdStr = "Invoke-WebClick -Selector `"$Selector`" -OuterHtml `"$outHtmlSafe`" -BoundingBox `"$bboxStr`""
        }
    }

    # 各要素を綺麗に分割したハッシュテーブルとして返却する。
    return @{
        command     = $cmdStr
        selector    = $Selector
        value       = $CurrentValue
        outerHtml   = $outHtmlSafe
        boundingBox = $BoundingBox
        isInsideFrame = $IsInsideFrame
    }
}

# ==============================================================================
# --- メインオーケストレーター ---
# 要素情報をもとに、DOM逆引き・UIA照合・フォールバック等のルート分岐を制御する。
# ==============================================================================
function Process-RecordAction {
    param(
        [bool]$IsGetNameMode,
        [bool]$UsePointReverse = $false
    )

    $func = $MyInvocation.MyCommand.Name

    try {
        # 1. OSレベルでのUIA要素取得 (対象要素の特定)： ピッカーによる操作対象を特定する。
        $elementInfo = Get-UiaTargetInfoFromCursor
        if (-not $elementInfo) { return $null }

        $targetElement     = $elementInfo.TargetElement
        $stepWindowName    = $elementInfo.WindowName
        $targetName        = $targetElement.Current.Name
        $targetId          = $targetElement.Current.AutomationId
        
        $pName = $targetElement.Current.ControlType.ProgrammaticName
        $targetControlType = if ($pName) { $pName.Replace("ControlType.", "") } else { "Unknown" }

#●         # デバッグ/診断用: 開発時コード（現在不使用）
#●         Test-WebSelectorResolution -UiaElement $targetElement | Out-Null

        # 2. UIAから現在の入力値（ValuePattern）を物理的に取得： DOM側で保持されていない値の奪取。
        $currentValue = ""
        $valuePattern = $null
        if ($targetElement.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$valuePattern)) {
            $currentValue = $valuePattern.Current.Value -replace "`r`n|`n|`r", " "
        }

        # ルートA： 物理座標からのDOM逆引き。
        if ($UsePointReverse -and ($stepWindowName -match "RPA Browser" -or $stepWindowName -match "Edge" -or $stepWindowName -match "Chrome")) {
            $rect = $targetElement.Current.BoundingRectangle
            $pt = New-Object System.Windows.Point([int]($rect.Left + ($rect.Width / 2)), [int]($rect.Top + ($rect.Height / 2)))
            $resObj = Invoke-BrowserPointReverse -ScreenPoint $pt -TargetControlType $targetControlType -CurrentValue $currentValue
            
            if ($resObj) { 
                $resObj.command = "[PointReverse] " + $resObj.command
                return "[RESULT]" + ($resObj | ConvertTo-Json -Compress)
            }
        }

        # ルートB： UIAメタデータからのDOM逆引き。
        $resObj = Invoke-UiaDomResolution -TargetName $targetName -TargetControlType $targetControlType -TargetId $targetId -TargetElement $targetElement -CurrentValue $currentValue
        if ($resObj) { 
            $resObj.command = "[UIA-DOM] " + $resObj.command
            return "[RESULT]" + ($resObj | ConvertTo-Json -Compress)
        }

        # ルートC： デスクトップ要素へのフォールバック。
        $resObj = Build-FallbackStepCode -WindowName $stepWindowName -TargetName $targetName -ControlType $targetControlType -CurrentValue $currentValue -IsGetNameMode $IsGetNameMode
        $resObj.command = "[Desktop-UIA] " + $resObj.command
        return "[RESULT]" + ($resObj | ConvertTo-Json -Compress)

    } catch {
        throw (New-EngineException -Func $func -Type "内部エラー" -Message "記録アクション生成中にエラーが発生しました" -Details $_.Exception.Message)
    }
}

# ------------------------------------------------------------------------------

# --- 最小単位要素の取得 ---
# 巨大コンテナ(Window等)を除外し、座標から最小単位のUIA要素を取得する。
function Get-UiaElementFromPoint {
    $func = $MyInvocation.MyCommand.Name

    try {
        # 現在のマウスカーソルのOS物理座標を取得し、Pointオブジェクトを生成する。
        $point = New-Object System.Windows.Point([System.Windows.Forms.Cursor]::Position.X, [System.Windows.Forms.Cursor]::Position.Y)
        $maxRetries = 15
        $waitMs = 200
        
        # 透明なオーバーレイや描画遅延を考慮し、一定回数リトライしながら要素を取得する。
        for ($retry = 1; $retry -le $maxRetries; $retry++) {
            # 指定座標にある最前面のUIA要素を取得する。
            $element = [System.Windows.Automation.AutomationElement]::FromPoint($point)
            # 取得した要素が操作対象として不適切な巨大コンテナ(Document, Pane, Window)でないか判定する。            
            if ($element -and 
                $element.Current.ControlType -ne [System.Windows.Automation.ControlType]::Document -and
                $element.Current.ControlType -ne [System.Windows.Automation.ControlType]::Pane -and
                $element.Current.ControlType -ne [System.Windows.Automation.ControlType]::Window -and
                $element.Current.AutomationId -ne "__next") { 
                
                return $element 
            }
            Start-Sleep -Milliseconds $waitMs
        }
        return $element
    } catch {
        throw (New-EngineException -Func $func -Type "GetElementError" -Message "座標からのUIA要素取得に失敗しました" -Details $_.Exception.Message)
    }
}

# --- 所属ウィンドウの特定 ---
# 要素の親階層を遡り、所属するウィンドウのUIA要素を特定する。
function Get-UiaWindowFromElement {
    param( [System.Windows.Automation.AutomationElement]$TargetElement )

    $func = $MyInvocation.MyCommand.Name

    try {
        # UIAのTreeWalkerを初期化し、コントロールとビューのみを対象としてツリーを探索可能にする。
        $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
        $current = $TargetElement
        # 要素がnullになるか、タイプがWindow(ウィンドウそのもの)に到達するまでループする。
        while ($current -ne $null -and $current.Current.ControlType -ne [System.Windows.Automation.ControlType]::Window) {
            # ひとつ上の親要素へ移動する。
            $current = $walker.GetParent($current)
        }
        return $current
    } catch {
        throw (New-EngineException -Func $func -Type "GetWindowError" -Message "親ウィンドウの特定に失敗しました" -Details $_.Exception.Message)
    }
}

# --- ウィンドウタイトルの正規化 ---
# ブラウザ特有の動的文字、ページ数等の揺れを吸収・正規化する。
function Format-UiaWindowName {
    param( [string]$RawWindowName )

    if ([string]::IsNullOrWhiteSpace($RawWindowName)) { return "" }
    return $RawWindowName `
        -replace "\s*および他\s*\d+\s*ページ.*$", "" `
        -replace "\s*-\s*(個人|仕事|InPrivate)(?=\s*-).*$", "" `
        -replace "\s*-[^-]+$", ""
}

# --- ターゲット要素のヒューリスティック探索 ---
# 直上の親要素も含めたスコアリング評価を行い、最適なUIA要素を探索する。
function Get-UiaTargetInfoFromCursor {
    $func = $MyInvocation.MyCommand.Name

    try {
        $targetElement = Get-UiaElementFromPoint
        if (-not $targetElement) { return $null }

        $walker = [System.Windows.Automation.TreeWalker]::ControlViewWalker
        $current = $targetElement

        # Window にぶつかるまで親を遡り(最大20階層)、最もRPA操作に適した要素を評価する。
        $bestCandidate = $null
        $bestScore = -1
        $depth = 0

        # 無限ループ防止のため最大20階層まで親を遡るループを実行する。
        while ($current -ne $null -and $depth -lt 20) {
            # 現在評価中の要素のプロパティを取得する。
            $name = $current.Current.Name
            $id   = $current.Current.AutomationId
            $type = $current.Current.ControlType.ProgrammaticName.Replace("ControlType.","")

            # 巨大なコンテナ要素(Window, Document, Pane)に到達した場合は探索を打ち切る。
            if ($current.Current.ControlType -eq [System.Windows.Automation.ControlType]::Window -or
                $current.Current.ControlType -eq [System.Windows.Automation.ControlType]::Document -or
                $current.Current.ControlType -eq [System.Windows.Automation.ControlType]::Pane -or
                $id -eq "__next") {
                break
            }

            $score = 0
            # NameやAutomationIdが存在すれば、セレクタとしての信頼性が高いため加点する。
            if (-not [string]::IsNullOrWhiteSpace($name)) { $score += 50 }
            if (-not [string]::IsNullOrWhiteSpace($id))   { $score += 40 }
            
            # 入力・操作系のコントロールタイプは優遇する。
            if ($type -in @("Edit","ComboBox","Button","Hyperlink","CheckBox","RadioButton")) {
                $score += 30
            }
            # 深すぎる階層は減点する（より直下の子要素を優先する）。
            $score -= ($depth * 2)

            # スコアがこれまでの最高値を上回ったら、候補を更新する。
            if ($score -gt $bestScore) {
                $bestScore = $score
                $bestCandidate = $current
            }
            # 次のループに向けて親要素へ移動し、階層カウンタを進める。
            $current = $walker.GetParent($current)
            $depth++
        }

        if ($bestCandidate) {
            $targetElement = $bestCandidate
        }

        $targetWindow = Get-UiaWindowFromElement -TargetElement $targetElement
        $windowName = if ($targetWindow) { Format-UiaWindowName -RawWindowName $targetWindow.Current.Name } else { "" }

        return @{ 
            TargetElement = $targetElement; 
            TargetWindow  = $targetWindow; 
            WindowName    = $windowName 
        }
    } catch {
        throw (New-EngineException -Func $func -Type "GetTargetInfoError" -Message "UIA探索中にエラー" -Details $_.Exception.Message)
    }
}

# --- 実行コマンド文字列の組み立て ---
# UIA/Web用の実行コマンド文字列(PowerShellコード)を組み立てる。
function Build-RpaStepCode {
    param(
        [string]$Action, 
        [string]$WindowName, 
        [string]$TargetName, 
        [string]$ControlType, 
        [int]$Index, 
        [string]$ActionValue, 
        [string]$AnchorName, 
        [string]$AnchorDirection, 
        [int]$AnchorSteps
    )

    $func = $MyInvocation.MyCommand.Name

    try {
        # 文字列内のダブルクォートをエスケープ： VBAやPSで安全に実行するため処理する。
        $safeActionValue = $ActionValue -replace '"', '`"'
        $safeTargetName  = $TargetName -replace '"', '`"'
        $safeWindowName  = $WindowName -replace '"', '`"'
        $safeAnchorName  = $AnchorName -replace '"', '`"'

        if ($Action -eq "GetName") {
            # 取得アクションの場合、結果を変数($extractedValue)へ代入し、コンソールへ出力するコードを組み立てる。
            return "`$extractedValue = Execute-UIRpaStep -TargetWindowName `"$safeWindowName`" -TargetName `"$safeTargetName`" -TargetType `"$ControlType`" -Index $Index -Action `"$Action`" -ActionValue `"$safeActionValue`" -AnchorName `"$safeAnchorName`" -AnchorDirection `"$AnchorDirection`" -AnchorSteps $AnchorSteps`r`nWrite-Host `"  => 取得した値: `$extractedValue`" -ForegroundColor Cyan"
        } else {
            # クリックなどの標準アクションの場合、そのまま実行用コマンドを組み立てて返却する。
            return "Execute-UIRpaStep -TargetWindowName `"$safeWindowName`" -TargetName `"$safeTargetName`" -TargetType `"$ControlType`" -Index $Index -Action `"$Action`" -ActionValue `"$safeActionValue`" -AnchorName `"$safeAnchorName`" -AnchorDirection `"$AnchorDirection`" -AnchorSteps $AnchorSteps"
        }
    } catch {
        throw (New-EngineException -Func $func -Type "CodeBuildError" -Message "ステップコードの生成に失敗しました" -Details $_.Exception.Message)
    }
}

# ==============================================================================
# --- 指定座標からの要素取得検証 ---
# 指定座標からの要素取得を検証する（VBAエントリーポイント・テスト用）。
# ==============================================================================
function Test-WebPointReverse {
    param (
        [int]$ClientX,
        [int]$ClientY
    )

    $func = $MyInvocation.MyCommand.Name
    Write-DebugLog -Message "================================================================================" -Level Info
    Write-DebugLog -Message "[$func] TEST座標: Client(X=$ClientX, Y=$ClientY)" -Level Info

    try {
        # 共通JSロジックを呼び出す。
        $jsPoint = Get-CorePointReverseJs -ClientX $ClientX -ClientY $ClientY

        # 手動ピッカーと同じ安全な関数(Invoke-WebScript)を使用してJSを実行する。
        $resRaw = Invoke-WebScript -Js $jsPoint
        
        Write-DebugLog -Message "[$func] 解析結果: $resRaw" -Level Info
        
        return "[RESULT]$resRaw"

    } catch {
        throw (New-EngineException -Func $func -Type "テスト実行エラー" -Message "テスト用JSの評価中にエラーが発生しました" -Details $_.Exception.Message)
    }
}

# ==============================================================================
# --- デバッグ/診断用 ---: 開発時コード（現在不使用）
# セレクタ解決能力のテスト出力を実行する。
# ==============================================================================
function Test-WebSelectorResolution {
    param([System.Windows.Automation.AutomationElement]$UiaElement)

    # UIA情報の取得
    $uiaName     = $UiaElement.Current.Name
    $uiaType     = $UiaElement.Current.ControlType.ProgrammaticName
    $uiaId       = $UiaElement.Current.AutomationId
    $uiaHelpText = $UiaElement.Current.HelpText

    Write-DebugLog -Message "[TEST] UIA Info: Name=$uiaName, Type=$uiaType, Id=$uiaId, HelpText=$uiaHelpText" -Level Info

    # [TEST:1] 位置ベースマッチング
    $selectorPos = Get-WebSelectorFromUia -UiaName $uiaName -UiaType $uiaType -UiaId $uiaId -UiaHelpText $uiaHelpText -UiaElement $UiaElement

    if ($selectorPos) { Write-DebugLog -Message "[TEST:1] 位置ベースUIA→DOM/ [OK] → Selector=$selectorPos" -Level Info }
    else { Write-DebugLog -Message "[TEST:1] 位置ベースUIA→DOM/ [NG] → Fuzzy へフォールバック" -Level Warning }

    # [TEST:2] 曖昧(Fuzzy)検索マッチング
    $selectorFuzzy = Get-WebSelectorFromUiaFuzzy -UiaName $uiaName -UiaType $uiaType -UiaId $uiaId -UiaHelpText $uiaHelpText

    if ($selectorFuzzy) { Write-DebugLog -Message "[TEST:2] Fuzzy(曖昧検索).../ [OK] → Selector=$selectorFuzzy" -Level Info }
    else { Write-DebugLog -Message "[TEST:2] Fuzzy(曖昧検索).../ [NG] Fuzzy でも一致なし" -Level Warning }

    # [TEST:3] 最終判定結果
    if ($selectorPos) {
        Write-DebugLog -Message "[TEST:3] 最終判定/ [RESULT] 位置マッチング → $selectorPos" -Level Info
        return $selectorPos
    } elseif ($selectorFuzzy) {
        Write-DebugLog -Message "[TEST:3] 最終判定/ [RESULT] Fuzzy(曖昧検索)→ $selectorFuzzy" -Level Info
        return $selectorFuzzy
    } else {
        Write-DebugLog -Message "[TEST:3] 最終判定/ [RESULT] UIA フォールバックへ移行" -Level Warning
        return $null
    }
}
