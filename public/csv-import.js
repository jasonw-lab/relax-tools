/**
 * CSVインポート機能（リボンコマンド用）
 * commands.htmlから読み込まれる
 */

/* global Excel */

// 列番号を列名に変換（0-based）
function getColumnName(colIndex) {
  let result = '';
  colIndex++; // 1-based に変換
  while (colIndex > 0) {
    colIndex--;
    result = String.fromCharCode(65 + (colIndex % 26)) + result;
    colIndex = Math.floor(colIndex / 26);
  }
  return result;
}

// CSVインポート機能（リボンコマンド用）
async function importCsvFromRibbon() {
  try {
    console.log('CSVインポート開始（リボンから）');
    
    // ファイル入力要素を作成
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.csv,text/csv';
    fileInput.style.display = 'none';
    
    // DOMに追加
    document.body.appendChild(fileInput);
    
    // ファイル選択を待つ
    const file = await new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        if (document.body.contains(fileInput)) {
          document.body.removeChild(fileInput);
        }
        reject(new Error('ファイル選択がタイムアウトしました'));
      }, 30000);
      
      const cleanup = () => {
        clearTimeout(timeout);
        if (document.body.contains(fileInput)) {
          document.body.removeChild(fileInput);
        }
      };
      
      fileInput.onchange = (e) => {
        cleanup();
        const target = e.target;
        const file = target.files?.[0];
        if (file) {
          console.log('ファイルが選択されました:', file.name);
          resolve(file);
        } else {
          reject(new Error('ファイルが選択されませんでした'));
        }
      };
      
      fileInput.addEventListener('cancel', () => {
        cleanup();
        reject(new Error('ファイル選択がキャンセルされました'));
      });
      
      setTimeout(() => {
        try {
          fileInput.click();
          console.log('ファイル選択ダイアログを開きました');
        } catch (clickError) {
          cleanup();
          console.error('ファイル選択ダイアログを開けませんでした:', clickError);
          reject(new Error('ファイル選択ダイアログを開けませんでした: ' + clickError.message));
        }
      }, 100);
    });
    
    // ファイルを読み込み
    console.log('ファイルを読み込み中...');
    const arrayBuffer = await file.arrayBuffer();
    const uint8Array = new Uint8Array(arrayBuffer);
    console.log('ファイルサイズ:', uint8Array.length, 'bytes');
    
    // エンコーディング判定
    let text;
    let encoding;
    
    // UTF-8 BOM判定
    if (uint8Array.length >= 3 && uint8Array[0] === 0xEF && uint8Array[1] === 0xBB && uint8Array[2] === 0xBF) {
      console.log('UTF-8 BOMを検出');
      text = new TextDecoder('utf-8').decode(uint8Array.slice(3));
      encoding = 'UTF-8';
    } else {
      // UTF-8として試行
      try {
        console.log('UTF-8としてデコードを試行');
        text = new TextDecoder('utf-8', { fatal: true }).decode(uint8Array);
        encoding = 'UTF-8';
        console.log('UTF-8としてデコード成功');
      } catch (e) {
        console.log('UTF-8デコード失敗、SJISとして試行:', e);
        // SJISとしてデコード（encoding-japaneseを使用）
        try {
          // 動的インポート（ESモジュールとして読み込む場合）
          const Encoding = await import('/node_modules/encoding-japanese/encoding.js');
          const result = Encoding.default.Convert(uint8Array, {
            to: 'UNICODE',
            from: 'SJIS',
            type: 'string',
          });
          if (typeof result === 'string') {
            text = result;
          } else {
            text = Encoding.default.codeToString(result);
          }
          encoding = 'SJIS';
          console.log('SJISとしてデコード成功');
        } catch (sjisError) {
          console.error('SJISデコードも失敗:', sjisError);
          // フォールバック: UTF-8として強制的にデコード
          text = new TextDecoder('utf-8', { fatal: false }).decode(uint8Array);
          encoding = 'UTF-8 (強制)';
          console.warn('エンコーディング判定に失敗、UTF-8として強制デコード');
        }
      }
    }
    
    console.log('デコード完了、文字数:', text.length);
    
    // CSVをパース（簡易実装：改行で分割）
    console.log('CSVをパース中...');
    const lines = text.split(/\r?\n/);
    const rows = [];
    for (const line of lines) {
      if (line.trim() === '') continue; // 空行をスキップ
      // 簡易CSVパース（カンマで分割、ダブルクォート処理は省略）
      const cells = line.split(',').map(cell => cell.trim().replace(/^"|"$/g, ''));
      rows.push(cells);
    }
    
    console.log('パース完了、行数:', rows.length);
    
    if (rows.length === 0) {
      throw new Error('CSVファイルが空です');
    }
    
    // 列数の計算
    const colCount = Math.max(...rows.map(row => row.length), 0);
    console.log('列数:', colCount);
    
    // Excelに書き込み
    console.log('Excelに書き込み開始');
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      const rowCount = rows.length;
      const rangeAddress = `A1:${getColumnName(colCount - 1)}${rowCount}`;
      console.log('範囲アドレス:', rangeAddress);
      
      const range = sheet.getRange(rangeAddress);
      
      // データを2次元配列に変換
      const values = [];
      for (const row of rows) {
        const excelRow = [];
        for (let i = 0; i < colCount; i++) {
          excelRow.push(row[i] || '');
        }
        values.push(excelRow);
      }
      
      console.log('データを設定、行数:', values.length, '列数:', colCount);
      range.values = values;
      
      // テーブル化
      console.log('テーブルを作成中...');
      const table = sheet.tables.add(rangeAddress, true);
      table.name = `Table_${Date.now()}`;
      table.style = 'TableStyleMedium2';
      
      await context.sync();
      console.log('Excelへの書き込み完了');
      console.log(`CSVインポート完了 (${encoding}, ${rowCount}行 x ${colCount}列)`);
    });
  } catch (error) {
    console.error('CSVインポートエラー:', error);
    const errorMessage = error instanceof Error ? error.message : String(error);
    console.error('エラーメッセージ:', errorMessage);
    // Officeアドイン環境では alert が使えないため、エラーを再スロー
    throw error;
  }
}

// グローバル関数としてエクスポート
if (typeof window !== 'undefined') {
  window.importCsv = importCsvFromRibbon;
}

