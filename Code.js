function scrapeWebPageTreeHorizontal() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const url = sheet.getRange("B1").getValue();  // B1セルにURL

    if (!url) {
        SpreadsheetApp.getUi().alert("B1セルにURLを入力してください。");
        return;
    }

    const response = UrlFetchApp.fetch(url);
    const html = response.getContentText();

    // <!DOCTYPE> を除去してXmlServiceでパース
    const cleanHtml = html.replace(/<!DOCTYPE[^>]*>/i, "");

    // 不正な & を &amp; に変換（既存の &amp; はそのまま）
    cleanHtml = cleanHtml.replace(/&(?![a-zA-Z0-9#]+;)/g, "&amp;");

    const document = XmlService.parse(cleanHtml);
    const root = document.getRootElement();
    let rows = [];

    // 再帰的にツリーを探索し、パスを格納
    function traverse(element, path) {
        const tag = element.getName();
        const href = element.getAttribute("href") ? element.getAttribute("href").getValue() : "";
        const classAttr = element.getAttribute("class") ? element.getAttribute("class").getValue() : "";
        const idAttr = element.getAttribute("id") ? element.getAttribute("id").getValue() : "";
        const text = element.getText().trim();

        // 新しいパスを作成
        const newPath = path.concat([[tag, href, classAttr, idAttr, text]]);

        // このノードまでのパスを一行として記録
        rows.push(newPath);

        // 子要素を探索
        const children = element.getChildren();
        if (children.length > 0) {
            children.forEach(child => traverse(child, newPath));
        }
    }

    traverse(root, []);

    // 最大階層を取得
    let maxDepth = 0;
    rows.forEach(path => {
        if (path.length > maxDepth) {
            maxDepth = path.length;
        }
    });

    // ヘッダ行を作成
    let header = [];
    for (let i = 1; i <= maxDepth; i++) {
        header = header.concat([`tag${i}`, `href${i}`, `class${i}`, `id${i}`, `text${i}`]);
    }

    // データを整形（足りない階層は空セルで埋める）
    let data = [header];
    rows.forEach(path => {
        let row = [];
        for (let i = 0; i < maxDepth; i++) {
            if (path[i]) {
                row = row.concat(path[i]);
            } else {
                row = row.concat(["", "", "", "", ""]);
            }
        }
        data.push(row);
    });

    // 出力（A2セルから）
    sheet.getRange(2, 1, data.length, data[0].length).setValues(data);
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
