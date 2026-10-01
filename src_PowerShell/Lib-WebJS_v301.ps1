# ------------------------------------------------------------------------------
<#
  [WebJS 共通ユーティリティモジュール]： WebView2初期化時に各ページやiframeへ自動注入されるグローバル関数群を定義する。
  RPA操作特有の課題（Shadow DOMの貫通、iframeの透過探索、厳密な可視性判定など）を解決し、
  OSネイティブ操作(UIA)とDOMをシームレスに繋ぐための基盤APIとして機能する。
#>
# ------------------------------------------------------------------------------

# ==============================================================================
# [基本JSユーティリティ関数の定義]： Shadow DOM貫通やiframe透過探索など、DOM操作の根幹となる関数群を定義する。
$global:ENGINE_JS_UTILS = @"

// --- Shadow DOM 貫通検索 ---
// 標準のquerySelectorでは取得できないShadow DOMの内部を再帰的に探索する。
function deepQuerySelector(selector, root) {
    // 探索の起点となるルートノードを設定する。引数がない場合はdocumentとする。
    root = root || document;
    // 現在のルート内で対象要素を通常のDOM探索で探す。
    let res = root.querySelector(selector);
    if (res) return res;
    // 現在のルート内のすべての要素をフラットに取得する。
    const els = root.querySelectorAll('*');
    for (let i = 0; i < els.length; i++) {
        // 要素がShadow DOMホスト(shadowRootを保持している)であるか判定する。
        if (els[i].shadowRoot) {
            // Shadow DOMの内部を新たなルートとして再帰的に探索を行う。
            res = deepQuerySelector(selector, els[i].shadowRoot);
            if (res) return res;
        }
    }
    return null;
}

// --- iframe透過探索 ---
// 同一オリジンのiframeを再帰的に探索し、対象が見つかれば返却する。
// クロスドメイン(CORS)制約によるアクセス拒否エラーは安全に無視し、探索を継続する。
function utilFindInFrames(win, checkFn) {
    try {
        const res = checkFn(win);
        if (res !== null && res !== false && res !== undefined) return res;
    } catch(e) {}
    for (let i = 0; i < win.frames.length; i++) {
        try {
            const found = utilFindInFrames(win.frames[i], checkFn);
            if (found !== null && found !== false && found !== undefined) return found;
        } catch(e) {}
    }
    return null;
}
"@

# ==============================================================================
# [デバッグ・拡張JSユーティリティ関数の定義]： 要素の可視性判定やセレクタ逆生成など、高度な解析用関数群を定義する。
$global:ENGINE_DEBUG_JS_UTILS = @"

// --- 内部トレーサー（ログ記録機構） ---
// DOM操作の過程や可視性判定の理由をスタックし、VBAやPowerShell側へ連携する。
const domTracker = {
    logs: [],
    log: function(msg) { this.logs.push(msg); },
    clear: function() { this.logs = []; }
};

// --- CSSセレクタ用特殊文字のエスケープ ---
// CSSセレクタとして無効な特殊文字をエスケープ処理する。
function cssEscape(str) {
    if (typeof CSS !== 'undefined' && typeof CSS.escape === 'function') return CSS.escape(str);
    return String(str).replace(/([^\x20-\x7E]|[ !"#$%&'()*+,./:;<=>?@[\\\]^`{|}~])/g, function (ch) { return '\\' + ch; });
}

// --- 要素の厳密な可視性判定 ---
// RPAで誤動作の原因となる「透明化された要素」「サイズゼロの要素」「裏に隠れた要素」を、人間の視覚に近い基準で正確に除外するための安全装置として機能する。
function isVisible(el, { checkOpacity = true, minOpacity = 0.01, silent = false } = {}) {
    // 要素が存在しない、またはDOMノード(NodeType=1)でない場合は非表示とみなす。
    if (!el || el.nodeType !== 1) {
        domTracker.log('isVisible: 失敗 - 要素が null または無効なノードです。');
        return false;
    }
    // 要素に適用されている最終的なCSSスタイルをOS(ブラウザ)から取得する。
    const style = window.getComputedStyle(el);
    if (!style) return false;
    
    // displayやvisibilityによる完全な非表示設定を検知する。
    if (style.display === 'none' || style.visibility === 'hidden' || style.visibility === 'collapse') return false;
    // hidden属性やaria-hidden属性によるセマンティックな非表示設定を検知する。
    if (el.hasAttribute('hidden') || el.getAttribute('aria-hidden') === 'true') return false;
    
    if (checkOpacity) {
        // opacity(透明度)を浮動小数点数として解析する。
        const opacity = parseFloat(style.opacity);
        // 指定された閾値以下であれば「視覚的に見えない(非表示)」と判定する。
        if (!Number.isNaN(opacity) && opacity <= minOpacity) return false;
    }
    
    // 要素がレンダリングツリー上にレイアウト領域を持っているか確認する。親要素ごと display: none されているケース等、視覚的に領域を持たない要素を弾く。
    const rects = el.getClientRects();
    if (!rects || rects.length === 0) return false;
    
    // 要素の絶対座標と寸法(BoundingBox)を取得する。
    const rect = el.getBoundingClientRect();
    // 幅または高さが0以下の場合は、画面上に実体がないため非表示と判定する。
    if (rect.width <= 0 || rect.height <= 0) return false;
    
    if (!silent) domTracker.log('isVisible: 要素は可視状態です (幅: ' + Math.round(rect.width) + ', 高さ: ' + Math.round(rect.height) + ')');
    return true;
}

// --- XPathを利用した複数DOM要素の一括取得 ---
// document.evaluateを利用し、XPathに合致するすべてのノードを配列として返却する。
function getElementsByXPath(xpath, root = document) {
    try {
        const result = document.evaluate(xpath, root, null, XPathResult.ORDERED_NODE_SNAPSHOT_TYPE, null);
        const nodes = [];
        for (let i = 0; i < result.snapshotLength; i++) {
            const node = result.snapshotItem(i);
            if (node && node.nodeType === 1) nodes.push(node);
        }
        return nodes;
    } catch (e) {
        return [];
    }
}

// --- 要素から一意に特定可能な強固なCSSセレクタの逆生成 ---
// 指定された要素から、DOM上で一意となるCSSセレクタを逆算して構築する。
function generateCssSelector(el, root = document) {
    if (!el || el.nodeType !== 1) throw new Error('element must be an Element node.');
    // htmlやbodyは一意性が自明であるため、そのまま返す。
    if (el === document.documentElement) return 'html';
    if (el === document.body) return 'body';

    // 優先度1位（id属性）： 存在すればページ内で最強の一意性を持つため優先する。
    const id = el.getAttribute && el.getAttribute('id');
    if (id) {
        // CSSセレクタ用にエスケープ処理を施したIDセレクタを構築する。
        const sel = '#' + cssEscape(id);
        // 構築したIDセレクタで検索し、自身が正確にヒットするか検証する。
        try { if ((root.querySelector(sel) === el)) return sel; } catch (e) {}
    }

    // 優先度2位（意味を持つ特定属性）： ボタンの役割やシステムが付与した名前などを利用する。
    const ATTR_CANDIDATES = ['name', 'type', 'title', 'placeholder', 'aria-label', 'role'];
    for (const attr of ATTR_CANDIDATES) {
        // 候補属性の値を要素から取得する。
        const val = el.getAttribute && el.getAttribute(attr);
        if (val) {
            // 属性セレクタ（[name="xxx"]等）を構築する。
            const sel = '[' + attr + '="' + cssEscape(val) + '"]';
            // この属性だけで一意に特定できるか（1件だけヒットし、それが自身か）を確認する。
            try {
                // 構築したセレクタでドキュメント全体を検索する。
                const found = root.querySelectorAll(sel);
                // 結果が1件のみで、かつ対象要素と完全一致すれば採用する。
                if (found.length === 1 && found[0] === el) return sel;
            } catch (e) {}
        }
    }

    // 優先度3位（カスタムデータ属性）： モダンWebフレームワークがよく使う data-* 属性を利用する。
    for (const a of Array.from(el.attributes || []).filter(at => at.name.startsWith('data-'))) {
        const sel = '[' + a.name + '="' + cssEscape(a.value) + '"]';
        try {
            const found = root.querySelectorAll(sel);
            if (found.length === 1 && found[0] === el) return sel;
        } catch (e) {}
    }

    // 優先度4位（DOMツリーの階層構造）： タグ名やクラス、兄弟要素の順番などを組み合わせる。
    const parts = [];
    let node = el;
    const rootElement = (root === document) ? document.documentElement : root;
    
    // 要素から親へ向かって階層を遡りながらセレクタを組み立てる。
    while (node && node.nodeType === 1 && node !== rootElement) {
        let part = node.tagName.toLowerCase();
        // クラス名があれば最大3つまで追加して絞り込みの精度を高める。
        if (node.classList && node.classList.length > 0) {
            const classes = Array.from(node.classList).slice(0, 3).map(c => '.' + cssEscape(c)).join('');
            const withClass = part + classes;
            try {
                const found = root.querySelectorAll(withClass);
                // この時点で一意になれば、これ以上親を遡る必要はないため即採用する。
                if (found.length === 1 && found[0] === el) {
                    parts.unshift(withClass);
                    break;
                }
            } catch (e) {}
            if (classes.length < 100) part += classes;
        }
        
        // 兄弟要素の中での順番（nth-of-type）を計算する。
        // 要素の親ノードを取得する。
        const parent = node.parentNode;
        // 親が存在しなければ、ツリーの頂点に達したとみなしループを抜ける。
        if (!parent) { parts.unshift(part); break; }
        
        const tagName = node.tagName;
        let index = 0;
        // 親ノードの子要素群をループし、同名タグの出現位置をカウントする。
        for (let i = 0; i < parent.children.length; i++) {
            if (parent.children[i].tagName === tagName) index++;
            // 自身に到達したらカウントを終了する。
            if (parent.children[i] === node) break;
        }
        // 同名のタグが複数あれば、何番目かを指定する。
        // :nth-of-type疑似クラスを付与し、同一階層内での一意性を高める。
        if (index > 0) part += ':nth-of-type(' + index + ')';
        
        parts.unshift(part);
        const candidate = parts.join(' > ');
        // 組み立てた階層セレクタで一意になったか確認する。
        try {
            const found = root.querySelectorAll(candidate);
            if (found.length === 1 && found[0] === el) return candidate;
        } catch (e) {}
        node = parent; 
    }
    
    // 最終チェック： 組み立てた階層の配列を子孫結合子( > )で結合し、一つのセレクタ文字列にする。
    const finalCandidate = parts.join(' > ');
    try {
        // ドキュメント全体から最終候補のセレクタで検索し、一意に特定できるか検証する。
        const found = root.querySelectorAll(finalCandidate);
        if (found.length === 1 && found[0] === el) return finalCandidate;
    } catch (e) {}
    
    // 最終手段（絶対パスによる強制生成）： 上記すべてで特定できない場合、ルートからの完全な階層パスを構築する。
    // 上記のすべてで特定できなかった場合、DOMツリーのルートからの完全な階層（nth-child等）を構築する。
    // 探索の起点を元の要素にリセットする。
    node = el;
    // 絶対パスの各階層を格納するための配列を初期化する。
    const fullParts = [];

    // ノードが存在し、かつ要素ノード(NodeType=1)である限り、親へ向かってループを継続する。
    while (node && node.nodeType === 1) {
        // 現在のノードのタグ名を小文字で取得する。
        let p = node.tagName.toLowerCase();
        // ID属性の有無を確認する。
        const id = node.getAttribute && node.getAttribute('id');
        if (id) {
            // IDが存在する場合、IDセレクタを採用して配列の先頭に追加し、そこで一意性が担保されるため探索を打ち切る。
            p = '#' + cssEscape(id);
            fullParts.unshift(p);
            break;
        }
        if (node.classList && node.classList.length) {
            // クラスが存在する場合、すべてのクラス名をエスケープして連結し、精度を高める。
            p += Array.from(node.classList).map(c => '.' + cssEscape(c)).join('');
        } else {
            // クラスが存在しない場合、親要素内でのインデックス番号を計算して位置を特定する。
            const parent = node.parentNode;
            if (parent) {
                // 兄弟要素群の中での自身のインデックス位置(1始まり)を取得する。
                const index = Array.prototype.indexOf.call(parent.children, node) + 1;
                // 疑似クラス :nth-child を付与して位置を確定させる。
                p += ':nth-child(' + index + ')';
            }
        }
        // 構築した現在の階層のパスを配列の先頭に追加する。
        fullParts.unshift(p);
        // 対象を親ノードへ移し、探索をさらに上へと進める。
        node = node.parentNode;
    }
    // 構築した絶対パス配列を子孫結合子( > )で結合して返却する。
    return fullParts.join(' > ');
}

// --- スクロールおよび可視状態確認を伴う安全なクリック実行 ---
// 要素を画面内にスクロールし、可視状態を確認した上で安全にクリックする。
function safeClick(el, { visible = true, tryScrollIntoView = true } = {}) {
    if (!el) throw new Error('要素が null または未定義です。');
    try {
        if (visible && !isVisible(el, { silent: true })) {
            domTracker.log('safeClick: スクロール前の時点で要素が非表示と判定されました。');
            throw new Error('スクロール前の時点で要素が非表示と判定されました。');
        }
        
        if (tryScrollIntoView) {
            domTracker.log('safeClick: 画面中央へのスクロールを試行します。');
            try { el.scrollIntoView({ block: 'center', inline: 'center', behavior: 'auto' }); }
            catch (e) { domTracker.log('safeClick: 警告 - スクロールに失敗しました (' + e.message + ')。'); }
        }
        
        if (visible && !isVisible(el, { silent: true })) {
            domTracker.log('safeClick: スクロール後に要素が他の要素の裏に隠れたか、非表示になりました。');
            throw new Error('スクロール後に要素が他の要素の裏に隠れたか、非表示になりました。');
        }
          
        domTracker.log('safeClick: 対象要素へのネイティブクリックを実行します！');
        el.click();
    } catch (err) {
        let selectorHint = null;
        try { selectorHint = generateCssSelector(el); } catch (e) { selectorHint = '(セレクタの生成不可)'; }
        domTracker.log('safeClick: エラー - ' + err.message);
        throw new Error('安全なクリックに失敗しました。理由: ' + err.message + ' / セレクタ候補: ' + selectorHint);
    }
}

// 実装した各関数をグローバルオブジェクト(window.DOMUtils)として公開する。
window.DOMUtils = {
    tracker: domTracker,
    cssEscape: cssEscape,
    isVisible: isVisible,
    getElementsByXPath: getElementsByXPath,
    generateCssSelector: generateCssSelector,
    safeClick: safeClick
};
"@
