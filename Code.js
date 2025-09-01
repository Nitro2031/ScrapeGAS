function scrapeWebPageTreeHorizontal() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const url = sheet.getRange("B1").getValue();
    if (!url) {
        SpreadsheetApp.getUi().alert("B1セルにURLを入力してください。");
        return;
    }

    const res = UrlFetchApp.fetch(url);
    const rawHtml = res.getContentText();

    // より堅牢なサニタイズ実行（XMLパーサ向け）
    const cleanHtml = sanitizeHtmlForXml(rawHtml);

    // パース（root でラップ済み）
    let document;
    try {
        document = XmlService.parse(cleanHtml);
    } catch (e) {
        SpreadsheetApp.getUi().alert("XMLパースに失敗しました。\n" + e);
        throw e;
    }

    const root = document.getRootElement();

    // root の直下の子要素ごとにツリーを展開して rows を作る
    const rows = [];
    const children = root.getChildren();
    for (let i = 0; i < children.length; i++) {
        traverseToRows(children[i], [], rows);
    }

    // 最大階層
    const maxDepth = rows.reduce((m, p) => Math.max(m, p.length), 0);
    // ヘッダ作成
    let header = [];
    for (let i = 1; i <= maxDepth; i++) {
        header = header.concat([`tag${i}`, `href${i}`, `class${i}`, `id${i}`, `text${i}`]);
    }

    // データ整形
    const data = [header];
    rows.forEach(path => {
        const row = [];
        for (let i = 0; i < maxDepth; i++) {
            if (path[i]) row.push(...path[i]);
            else row.push("", "", "", "", "");
        }
        data.push(row);
    });

    // シートのサイズを確保（必要なら行/列を追加）
    const startRow = 2, startCol = 1;
    const needRows = data.length;
    const needCols = data[0].length;
    const curRows = sheet.getMaxRows();
    const curCols = sheet.getMaxColumns();
    const requiredTotalRows = startRow - 1 + needRows;
    const requiredTotalCols = startCol - 1 + needCols;
    if (curRows < requiredTotalRows) {
        sheet.insertRowsAfter(curRows, requiredTotalRows - curRows);
    }
    if (curCols < requiredTotalCols) {
        sheet.insertColumnsAfter(curCols, requiredTotalCols - curCols);
    }

    // 既存出力領域を一旦クリア（A2以降）
    sheet.getRange(startRow, startCol, sheet.getMaxRows() - startRow + 1, sheet.getMaxColumns()).clearContent();

    // 書き込み
    sheet.getRange(startRow, startCol, data.length, data[0].length).setValues(data);
}

/* ----------------- 助手関数: HTMLをXmlService.parse向けに整形 ----------------- */
function sanitizeHtmlForXml(html) {
    let s = html;

    // 1) <script> / <style> を取り出す（後でCDATAで戻す）
    const scriptStore = [];
    s = s.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, function (m) {
        const idx = scriptStore.length;
        scriptStore.push(m);
        return `__SCRIPT_PLACEHOLDER_${idx}__`;
    });
    const styleStore = [];
    s = s.replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, function (m) {
        const idx = styleStore.length;
        styleStore.push(m);
        return `__STYLE_PLACEHOLDER_${idx}__`;
    });

    // 2) コメントを削除（条件付きコメントも含む）
    s = s.replace(/<!--[\s\S]*?-->/g, "");

    // 3) DOCTYPE / XML宣言 を削除
    s = s.replace(/<!DOCTYPE[^>]*>/ig, "");
    s = s.replace(/<\?xml[^>]*\?>/ig, "");

    // 4) 属性内の '>' '<' を一時プレースホルダ化（タグマッチを壊さないため）
    s = s.replace(/("[^"]*"|'[^']*')/g, function (m) {
        const quote = m[0];
        const inner = m.slice(1, -1).replace(/>/g, "__GT__").replace(/</g, "__LT__");
        return quote + inner + quote;
    });

    // 5) タグごとに属性を正規化して再構築
    const voidTags = { area: 1, base: 1, br: 1, col: 1, embed: 1, hr: 1, img: 1, input: 1, link: 1, meta: 1, param: 1, source: 1, track: 1, wbr: 1 };
    s = s.replace(/<([a-zA-Z][\w:-]*)([^>]*)>/g, function (_m, tag, attrs) {
        attrs = attrs || "";
        const outAttrs = [];
        // attr を逐次パース（quoted / unquoted / boolean を扱う）
        const attrRegex = /([^\s=\/>]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'>]+)))?/g;
        let ma;
        while ((ma = attrRegex.exec(attrs)) !== null) {
            const name = ma[1];
            let val = (ma[2] !== undefined) ? ma[2] : (ma[3] !== undefined ? ma[3] : (ma[4] !== undefined ? ma[4] : null));
            if (val === null) {
                // boolean attribute -> attr="attr"
                val = name;
            }
            // プレースホルダを戻す
            val = val.replace(/__GT__/g, ">").replace(/__LT__/g, "<");
            // & をエスケープ（ただし既存の &name; 等は残す）
            val = val.replace(/&(?![a-zA-Z0-9#]+;)/g, "&amp;");
            // " があればエンティティ化
            val = val.replace(/"/g, "&quot;");
            // 特別扱い: crossorigin の簡易補完（値なしなら anonymous）
            if (/^crossorigin$/i.test(name) && (val === "" || val.toLowerCase() === "crossorigin")) {
                // if it was boolean style, val currently equals name -> set anonymous
                if (val.toLowerCase() === "crossorigin") val = "anonymous";
            }
            outAttrs.push(name + '="' + val + '"');
        }
        const attrString = outAttrs.length ? " " + outAttrs.join(" ") : "";
        if (voidTags[tag.toLowerCase()]) {
            return "<" + tag + attrString + "/>";
        } else {
            return "<" + tag + attrString + ">";
        }
    });

    // 6) 退避していた script/style を CDATA で戻す
    s = s.replace(/__SCRIPT_PLACEHOLDER_(\d+)__/g, function (_m, i) {
        const block = scriptStore[Number(i)];
        const mm = block.match(/<script([^>]*)>([\s\S]*?)<\/script>/i);
        const attrs = mm ? mm[1] : "";
        const body = mm ? mm[2] : "";
        return "<script" + attrs + "><![CDATA[" + body + "]]></script>";
    });
    s = s.replace(/__STYLE_PLACEHOLDER_(\d+)__/g, function (_m, i) {
        const block = styleStore[Number(i)];
        const mm = block.match(/<style([^>]*)>([\s\S]*?)<\/style>/i);
        const attrs = mm ? mm[1] : "";
        const body = mm ? mm[2] : "";
        return "<style" + attrs + "><![CDATA[" + body + "]]></style>";
    });

    // 7) 残った & をエスケープ（既にエンティティのものは除外）
    s = s.replace(/&(?![a-zA-Z0-9#]+;)/g, "&amp;");

    // 8) 必ず単一ルートにする
    return "<root>" + s + "</root>";
}

/* ------------- 助手関数: XML要素をたどってパス（[tag,href,class,id,text] の配列）を rows に詰める ------------- */
function traverseToRows(element, path, rows) {
    if (!element) return;
    const tag = element.getName();
    const href = element.getAttribute("href") ? element.getAttribute("href").getValue() : "";
    const classAttr = element.getAttribute("class") ? element.getAttribute("class").getValue() : "";
    const idAttr = element.getAttribute("id") ? element.getAttribute("id").getValue() : "";
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
