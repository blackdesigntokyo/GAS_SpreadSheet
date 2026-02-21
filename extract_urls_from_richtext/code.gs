function extractUrls() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  
  // データがない場合は終了
  if (lastRow === 0) return;
  
  // A列（1列目）の1行目から最終行までを取得
  const range = sheet.getRange(1, 1, lastRow, 1);
  const richTextValues = range.getRichTextValues();
  
  const results = [];
  let maxCols = 1; // 出力用の列数を管理
  
  for (let i = 0; i < richTextValues.length; i++) {
    const richText = richTextValues[i][0];
    const rowUrls = [];
    
    if (richText) {
      // セル内のテキストを要素ごとに分割して確認
      const runs = richText.getRuns();
      for (let j = 0; j < runs.length; j++) {
        const url = runs[j].getLinkUrl();
        // ＵＲＬが存在し、まだ追加されていない場合に追加（重複防止）
        if (url && rowUrls.indexOf(url) === -1) {
          rowUrls.push(url);
        }
      }
    }
    
    results.push(rowUrls);
    // 最大のＵＲＬ数に合わせて出力列数を更新
    if (rowUrls.length > maxCols) {
      maxCols = rowUrls.length;
    }
  }
  
  // スプレッドシートに書き込むために、配列の長さを統一
  for (let i = 0; i < results.length; i++) {
    while (results[i].length < maxCols) {
      results[i].push(""); // 空白で埋める
    }
  }
  
  // B列（2列目）から結果を出力
  const outputRange = sheet.getRange(1, 2, results.length, maxCols);
  // 出力先を一旦クリアしてから書き込み
  outputRange.clearContent();
  outputRange.setValues(results);
}
