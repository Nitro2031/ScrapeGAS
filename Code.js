function scrapeWebPageTreeHorizontal() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const url = sheet.getRange("B1").getValue();  // B1セルのURL
    if (!url) {
        SpreadsheetApp.getUi().alert("B1セルにURLを入力してください。");
        return;
    }

    const res = UrlFetchApp.fetch(url);
    const rawHtml = res.getContentText();

    // 1) HTML → XML風に整形（よくある崩れを修正）
    const cleanHtml = sanitizeHtmlForXml(rawHtml);

    // 2) XmlService でパース
    let document, root;
    try {
        document = XmlService.parse(cleanHtml);
        root = document.getRootElement();
    } catch (e) {
        // 解析に失敗した場合、どこで転んだかを見やすく表示
        SpreadsheetApp.getUi().alert("XMLパースに失敗しました。\n" + e);
        throw e;
    }

    // 3) ツリーをたどって「階層パス」を1行に
    const rows = [];
    traverseToRows(root, [], rows);

    // 4) 最大階層を測ってヘッダー作成（A2から）
    const maxDepth = rows.reduce((m, p) => Math.max(m, p.length), 0);
    let header = [];
    for (let i = 1; i <= maxDepth; i++) {
        header = header.concat([`tag${i}`, `href${i}`, `class${i}`, `id${i}`, `text${i}`]);
    }

    // 5) 行データを整形（足りない階層は空で埋める）
    const data = [header];
    rows.forEach(path => {
        let row = [];
        for (let i = 0; i < maxDepth; i++) {
            if (path[i]) row = row.concat(path[i]);
            else row = row.concat(["", "", "", "", ""]);
        }
        data.push(row);
    });

    // 6) 出力（A2セルから）
    const startRow = 2, startCol = 1;
    if (data.length && data[0].length) {
        // 既存の出力領域を一旦クリア（横幅は今回の列数ぶん）
        sheet.getRange(startRow, startCol, Math.max(1, sheet.getMaxRows() - startRow + 1), data[0].length).clearContent();
        sheet.getRange(startRow, startCol, data.length, data[0].length).setValues(data);
    }
}

/**
 * HTML文字列を、XmlService.parse() が読めるように“できるだけ”整形します。
 * - DOCTYPE除去
 * - 未エスケープ & をエスケープ
 * - 引用符なし属性値をクォート
 * - 値なし属性（boolean属性）を attr="attr" へ
 * - crossorigin が値なしの場合は crossorigin="anonymous" に
 * - VOIDタグを自閉（<link> → <link /> など）
 * - <script>/<style> 内容は CDATA で保護
 */
function sanitizeHtmlForXml(html) {
    let s = html;

    // --- 0) <script>/<style> を一時退避（中身を汚さないため） ---
    const scriptStore = [];
    s = s.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, (m) => {
        const idx = scriptStore.length;
        scriptStore.push(m);
        return `__SCRIPT_PLACEHOLDER_${idx}__`;
    });
    const styleStore = [];
    s = s.replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, (m) => {
        const idx = styleStore.length;
        styleStore.push(m);
        return `__STYLE_PLACEHOLDER_${idx}__`;
    });

    // 1) DOCTYPE除去
    s = s.replace(/<!DOCTYPE[^>]*>/ig, "");

    // 2) 未エスケープの & をエスケープ（&name; / &#123; / &#x1F; は残す）
    s = s.replace(/&(?![a-zA-Z0-9#]+;)/g, "&amp;");

    // 3) タグ内の属性を正規化
    s = s.replace(/<([a-zA-Z][\w:-]*)([^>]*)>/g, function (_m, tag, attrs) {
        let a = attrs || "";

        // 3-1) 引用符なしの属性値: key=value → key="value"
        a = a.replace(/(\s)([a-zA-Z_:][\w:.-]*)=([^\s"'=<>`]+)/g, '$1$2="$3"');

        // 3-2) 値なし（boolean）属性を attr="attr" に
        const booleanAttrs = [
            "async", "defer", "disabled", "checked", "selected", "autofocus", "autoplay",
            "controls", "loop", "muted", "playsinline", "hidden", "multiple", "readonly",
            "required", "scoped", "nomodule", "open", "download", "ismap", "reversed",
            "itemscope", "allowfullscreen", "formnovalidate", "novalidate", "inert"
        ];
        booleanAttrs.forEach(attr => {
            const re = new RegExp("(\\s)" + attr + "(?=(\\s|>|/))", "ig");
            a = a.replace(re, `$1${attr}="${attr}"`);
        });

        // 3-3) crossorigin が値なしなら既定で anonymous を付与
        a = a.replace(/(\s)crossorigin(?=(\s|>|\/))/ig, '$1crossorigin="anonymous"');

        // 3-4) 末尾の余分な空白を削除
        a = a.replace(/\s+$/, "");

        return `<${tag}${a}>`;
    });

    // 4) VOID要素を自閉に（XML準拠にする）
    s = s.replace(/<(area|base|br|col|embed|hr|img|input|link|meta|param|source|track|wbr)([^\/>]*?)>/gi, "<$1$2/>");

    // --- 5) 退避していた <script>/<style> を CDATA で戻す ---
    s = s.replace(/__SCRIPT_PLACEHOLDER_(\d+)__/g, (_m, i) => {
        const block = scriptStore[Number(i)];
        // 属性部は tag 正規化に任せるより、そのまま戻しつつ中身だけ CDATA 保護
        const m2 = block.match(/<script([^>]*)>([\s\S]*?)<\/script>/i);
        const attrs = m2 ? m2[1] : "";
        const body = m2 ? m2[2] : "";
        return `<script${attrs}><![CDATA[${body}]]></script>`;
    });
    s = s.replace(/__STYLE_PLACEHOLDER_(\d+)__/g, (_m, i) => {
        const block = styleStore[Number(i)];
        const m2 = block.match(/<style([^>]*)>([\s\S]*?)<\/style>/i);
        const attrs = m2 ? m2[1] : "";
        const body = m2 ? m2[2] : "";
        return `<style${attrs}><![CDATA[${body}]]></style>`;
    });

    return s;
}

/**
 * XML要素をたどって、各ノードまでの「パス」を rows に詰める
 * path の各要素は [tag, href, class, id, text]
 */
function traverseToRows(element, path, rows) {
    if (!element) return;

    const tag = element.getName();
    const href = element.getAttribute("href") ? element.getAttribute("href").getValue() : "";
    const classAttr = element.getAttribute("class") ? element.getAttribute("class").getValue() : "";
    const idAttr = element.getAttribute("id") ? element.getAttribute("id").getValue() : "";

    // script/style は text を空に（コードやCSSが巨大化するのを防ぐ）
    const text = (/^(script|style)$/i.test(tag)) ? "" : element.getText().trim();

    const newPath = path.concat([[tag, href, classAttr, idAttr, text]]);
    rows.push(newPath);

    const children = element.getChildren();
    for (let i = 0; i < children.length; i++) {
        traverseToRows(children[i], newPath, rows);
    }
}

/**
 * Google スプレッドシートのメニューに「📦 ScrapeWebPage」というカスタムメニューを追加する関数
 * @returns {void}
 */
function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('📦 ScrapeWebPage')
        .addItem('GetWebPage', 'scrapeWebPageTreeHorizontal')
        .addToUi();
}
