/**
 * CSV取込: 簡易CSVインポート
 * ファイル選択 → UTF-8/SJIS自動判定 → A1に書き込み → テーブル化
 */

/* global Excel */

/**
 * 簡易CSVインポート機能
 * ファイル選択ダイアログでCSVを選択し、A1から書き込んでテーブル化
 */
export async function importCsvQuick(): Promise<void> {
  try {
    console.log('CSVインポート開始');
    
    // ファイル入力要素を作成（非表示だが、DOMに追加する必要がある）
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.csv,text/csv';
    fileInput.style.display = 'none'; // 非表示にする
    
    // DOMに追加（Officeアドイン環境では必要）
    document.body.appendChild(fileInput);
    
    // ファイル選択を待つ
    console.log('ファイル選択ダイアログを開きます');
    const file = await new Promise<File>((resolve, reject) => {
      // タイムアウトを設定（30秒）
      const timeout = setTimeout(() => {
        document.body.removeChild(fileInput); // クリーンアップ
        reject(new Error('ファイル選択がタイムアウトしました'));
      }, 30000);
      
      const cleanup = () => {
        clearTimeout(timeout);
        if (document.body.contains(fileInput)) {
          document.body.removeChild(fileInput); // クリーンアップ
        }
      };
      
      fileInput.onchange = (e) => {
        cleanup();
        const target = e.target as HTMLInputElement;
        const file = target.files?.[0];
        if (file) {
          console.log('ファイルが選択されました:', file.name);
          resolve(file);
        } else {
          reject(new Error('ファイルが選択されませんでした'));
        }
      };
      
      // キャンセル処理（ただし、ブラウザによっては発火しない場合がある）
      fileInput.addEventListener('cancel', () => {
        cleanup();
        reject(new Error('ファイル選択がキャンセルされました'));
      });
      
      // 少し遅延を入れてからクリック（DOMに追加されたことを確実にする）
      setTimeout(() => {
        try {
          fileInput.click();
          console.log('ファイル選択ダイアログを開きました');
        } catch (clickError) {
          cleanup();
          console.error('ファイル選択ダイアログを開けませんでした:', clickError);
          reject(new Error('ファイル選択ダイアログを開けませんでした: ' + (clickError as Error).message));
        }
      }, 100);
    });
    
    // ファイルを読み込み
    console.log('ファイルを読み込み中...');
    const arrayBuffer = await file.arrayBuffer();
    const uint8Array = new Uint8Array(arrayBuffer);
    console.log('ファイルサイズ:', uint8Array.length, 'bytes');
    
    // エンコーディング判定（簡易版: 先頭2バイトでBOM判定、なければSJISとして試行）
    let text: string;
    let encoding: string;
    
    // UTF-8 BOM判定 (EF BB BF)
    if (uint8Array.length >= 3 && uint8Array[0] === 0xEF && uint8Array[1] === 0xBB && uint8Array[2] === 0xBF) {
      console.log('UTF-8 BOMを検出');
      text = new TextDecoder('utf-8').decode(uint8Array.slice(3));
      encoding = 'UTF-8';
    } else {
      // encoding-japanese を使用してSJIS判定
      // 簡易実装: まずUTF-8として試行、失敗したらSJIS
      try {
        console.log('UTF-8としてデコードを試行');
        text = new TextDecoder('utf-8', { fatal: true }).decode(uint8Array);
        encoding = 'UTF-8';
        console.log('UTF-8としてデコード成功');
      } catch (e) {
        console.log('UTF-8デコード失敗、SJISとして試行:', e);
        // SJISとしてデコード
        // encoding-japanese の Convert を使用
        try {
          const Encoding = (await import('encoding-japanese')).default;
          console.log('encoding-japaneseをインポート完了');
          
          // Convert 関数を使用してSJISからUNICODEに変換
          // type: 'string' オプションで直接文字列を取得できる
          const result = Encoding.Convert(uint8Array, {
            to: 'UNICODE',
            from: 'SJIS',
            type: 'string',
          });
          // type: 'string' がサポートされている場合は文字列、そうでなければコード配列
          if (typeof result === 'string') {
            text = result;
            console.log('SJISとしてデコード成功（文字列）');
          } else {
            // コード配列の場合は codeToString で変換
            text = Encoding.codeToString(result);
            console.log('SJISとしてデコード成功（コード配列→文字列）');
          }
          encoding = 'SJIS';
        } catch (sjisError) {
          console.error('SJISデコードも失敗:', sjisError);
          throw new Error(`文字エンコーディングの判定に失敗しました: ${sjisError}`);
        }
      }
    }
    
    console.log('デコード完了、文字数:', text.length);
    
    // Papa.parse でCSVをパース
    console.log('CSVをパース中...');
    const Papa = (await import('papaparse')).default;
    const parseResult = Papa.parse(text, {
      header: false,
      skipEmptyLines: false,
    }) as { data: string[][]; errors: any[]; meta: any };
    
    if (parseResult.errors.length > 0) {
      console.warn('CSVパース警告:', parseResult.errors);
    }
    
    const rows = parseResult.data as string[][];
    console.log('パース完了、行数:', rows.length);
    
    if (rows.length === 0) {
      console.warn('CSVファイルが空です');
      throw new Error('CSVファイルが空です');
    }
    
    // 列数の計算
    const colCount = Math.max(...rows.map((row: string[]) => row.length), 0);
    console.log('列数:', colCount);
    
    // Excelに書き込み
    console.log('Excelに書き込み開始');
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // A1からデータを書き込み
      const rowCount = rows.length;
      
      // 範囲のアドレス文字列を直接構築（range.addressを使わない）
      const rangeAddress = `A1:${getColumnName(colCount - 1)}${rowCount}`;
      console.log('範囲アドレス:', rangeAddress);
      
      // 範囲を準備
      const range = sheet.getRange(rangeAddress);
      
      // データを2次元配列に変換（空セルは null ではなく空文字列）
      const values: (string | number | boolean | null)[][] = [];
      for (const row of rows) {
        const excelRow: (string | number | boolean | null)[] = [];
        for (let i = 0; i < colCount; i++) {
          excelRow.push(row[i] || '');
        }
        values.push(excelRow);
      }
      
      console.log('データを設定、行数:', values.length, '列数:', colCount);
      range.values = values;
      
      // テーブル化（最初の行をヘッダーとして使用）
      // range.address を使わず、直接アドレス文字列を使用
      console.log('テーブルを作成中...');
      const table = sheet.tables.add(rangeAddress, true); // hasHeaders = true
      table.name = `Table_${Date.now()}`; // 一意な名前
      table.style = 'TableStyleMedium2'; // デフォルトスタイル
      
      await context.sync();
      console.log('Excelへの書き込み完了');
      console.log(`CSVインポート完了 (${encoding}, ${rowCount}行 x ${colCount}列)`);
    });
  } catch (error) {
    console.error('CSVインポートエラー:', error);
    const errorMessage = error instanceof Error ? error.message : String(error);
    const errorStack = error instanceof Error ? error.stack : '';
    console.error('エラースタック:', errorStack);
    // Officeアドイン環境では alert() が使えないため、エラーを throw して
    // App.tsx の handleFeatureClick で処理させる
    throw new Error(`CSVのインポートに失敗しました: ${errorMessage}`);
  }
}

/**
 * 列番号を列名に変換（0-based）
 * 例: 0 -> A, 25 -> Z, 26 -> AA
 */
function getColumnName(colIndex: number): string {
  let result = '';
  colIndex++; // 1-based に変換
  while (colIndex > 0) {
    colIndex--;
    result = String.fromCharCode(65 + (colIndex % 26)) + result;
    colIndex = Math.floor(colIndex / 26);
  }
  return result;
}

