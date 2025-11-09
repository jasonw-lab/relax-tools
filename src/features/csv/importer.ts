/**
 * CSV取込: 簡易CSVインポート（カード）
 * 複数ファイル選択 → 各ファイルごとに月を抽出 → 該当シートに書き込み
 */

/* global Excel */

/**
 * ファイル名から月を抽出する
 * ファイル名形式：`enaviYYMMDD(XXXX).csv` または `enaviYYYYMM(XXXX).csv`
 * 例：
 *   - `enavi202506(3034).csv` → "6"（6月）
 *   - `enavi202507(3034).csv` → "7"（7月）
 * @param fileName ファイル名
 * @returns 月（文字列）、抽出できない場合は null
 */
function extractMonthFromFileName(fileName: string): string | null {
  // ファイル名から拡張子を除去
  const nameWithoutExt = fileName.replace(/\.csv$/i, '');
  
  // パターン: enavi + 6桁の数字 + (XXXX)
  const pattern = /^enavi(\d{6})\(/;
  const match = nameWithoutExt.match(pattern);
  if (!match) {
    return null;
  }
  
  const digits = match[1];
  
  // 最初の2桁または4桁で形式を判定
  // YYYYMM形式: 最初の4桁が年（例: 2025 → 202506）
  // YYMMDD形式: 最初の2桁が年（例: 25 → 250615）
  
  // 最初の4桁が20以上なら YYYYMM 形式と判定
  const first4Digits = parseInt(digits.substring(0, 4), 10);
  if (first4Digits >= 2000) {
    // YYYYMM 形式: 最後の2桁が月
    const month = digits.substring(4);
    const monthNum = parseInt(month, 10);
    if (monthNum >= 1 && monthNum <= 12) {
      return monthNum.toString();
    }
  } else {
    // YYMMDD 形式: 3-4桁目が月
    const month = digits.substring(2, 4);
    const monthNum = parseInt(month, 10);
    if (monthNum >= 1 && monthNum <= 12) {
      return monthNum.toString();
    }
  }
  
  return null;
}

/**
 * CSVファイルを読み込んでテキストに変換
 * @param file ファイルオブジェクト
 * @returns テキストとエンコーディング
 */
async function readCsvFile(file: File): Promise<{ text: string; encoding: string }> {
  const arrayBuffer = await file.arrayBuffer();
  const uint8Array = new Uint8Array(arrayBuffer);
  
  let text: string;
  let encoding: string;
  
  // UTF-8 BOM判定 (EF BB BF)
  if (uint8Array.length >= 3 && uint8Array[0] === 0xEF && uint8Array[1] === 0xBB && uint8Array[2] === 0xBF) {
    text = new TextDecoder('utf-8').decode(uint8Array.slice(3));
    encoding = 'UTF-8';
  } else {
    try {
      text = new TextDecoder('utf-8', { fatal: true }).decode(uint8Array);
      encoding = 'UTF-8';
    } catch (e) {
      const Encoding = (await import('encoding-japanese')).default;
      const result = Encoding.Convert(uint8Array, {
        to: 'UNICODE',
        from: 'SJIS',
        type: 'string',
      });
      if (typeof result === 'string') {
        text = result;
      } else {
        text = Encoding.codeToString(result);
      }
      encoding = 'SJIS';
    }
  }
  
  return { text, encoding };
}

/**
 * CSV取込機能（カード）
 * 複数のCSVファイルを選択し、各ファイル名から月を抽出して該当シートにインポート
 */
export async function importCsvQuick(): Promise<void> {
  try {
    console.log('CSVインポート開始（カード）');
    
    // ファイル入力要素を作成（複数選択可能）
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.csv,text/csv';
    fileInput.multiple = true; // 複数選択を有効化
    fileInput.style.display = 'none';
    
    document.body.appendChild(fileInput);
    
    // ファイル選択を待つ
    const files = await new Promise<File[]>((resolve, reject) => {
      const timeout = setTimeout(() => {
        document.body.removeChild(fileInput);
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
        const target = e.target as HTMLInputElement;
        const files = target.files ? Array.from(target.files) : [];
        if (files.length > 0) {
          console.log(`${files.length}個のファイルが選択されました`);
          resolve(files);
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
        } catch (clickError) {
          cleanup();
          reject(new Error('ファイル選択ダイアログを開けませんでした: ' + (clickError as Error).message));
        }
      }, 100);
    });
    
    // 各シートごとの行ポインタ管理（シート名 -> 行番号）
    const sheetRowPtr: Record<string, number> = {};
    // 各シートのクリア済みフラグ（シート名 -> boolean）
    const sheetCleared: Record<string, boolean> = {};
    // 警告メッセージのリスト
    const warnings: string[] = [];
    
    // 各CSVファイルを処理
    for (const file of files) {
      try {
        // ファイル名から月を抽出
        const month = extractMonthFromFileName(file.name);
        if (!month) {
          warnings.push(`ファイル「${file.name}」: 月を抽出できませんでした`);
          continue;
        }
        
        // 該当シートの存在チェック
        let sheetExists = false;
        await Excel.run(async (context) => {
          try {
            const sheet = context.workbook.worksheets.getItem(month);
            sheet.load('name');
            await context.sync();
            sheetExists = true;
          } catch (error) {
            sheetExists = false;
          }
        });
        
        if (!sheetExists) {
          warnings.push(`ファイル「${file.name}」: 該当シート「${month}」が存在しません`);
          continue;
        }
        
        // 各シートにつき最初の書き込み前に1回だけクリア
        if (!sheetCleared[month]) {
          await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(month);
            const clearRange = sheet.getRange('A4:L200');
            // セル内容をクリア
            clearRange.clear(Excel.ClearApplyTo.contents);
            // 背景色もクリア（罫線は維持）
            clearRange.format.fill.clear();
            await context.sync();
          });
          sheetCleared[month] = true;
          // クリア後は行ポインタを4に初期化
          sheetRowPtr[month] = 4;
        }
        
        // 現在の行ポインタを取得（初期化されていない場合は4）
        let rowPtr = sheetRowPtr[month] || 4;
        
        // ファイルを読み込み
        const { text } = await readCsvFile(file);
        
        // Papa.parse でCSVをパース（skipEmptyLines: "greedy"）
        const Papa = (await import('papaparse')).default;
        const parseResult = Papa.parse(text, {
          header: false,
          skipEmptyLines: 'greedy' as any,
        }) as { data: string[][]; errors: any[]; meta: any };
        
        if (parseResult.errors.length > 0) {
          console.warn('CSVパース警告:', parseResult.errors);
        }
        
        const rows = parseResult.data as string[][];
        
        if (rows.length === 0) {
          warnings.push(`ファイル「${file.name}」: CSVファイルが空です`);
          continue;
        }
        
        // ヘッダー行を除外してデータ行だけ扱う（2行目以降）
        const dataRows = rows.length >= 2 ? rows.slice(1) : [];
        
        if (dataRows.length === 0) {
          warnings.push(`ファイル「${file.name}」: データ行がありません`);
          continue;
        }
        
        // 列数はA〜L（12列）に制限
        const MAX_COLS = 12;
        const dataRowsNormalized: (string | number | boolean | null)[][] = [];
        for (const row of dataRows) {
          const excelRow: (string | number | boolean | null)[] = [];
          for (let i = 0; i < MAX_COLS; i++) {
            excelRow.push(i < row.length ? (row[i] || '') : '');
          }
          dataRowsNormalized.push(excelRow);
        }
        
        // ファイル名を書き込む行
        const fileNameRow: (string | number | boolean | null)[] = [file.name];
        for (let i = 1; i < MAX_COLS; i++) {
          fileNameRow.push('');
        }
        
        // 書き込み上限チェック（L200 = 行200）
        const MAX_ROW = 200;
        const fileNameRowNum = rowPtr;
        const dataStartRow = rowPtr + 1;
        const dataEndRow = dataStartRow + dataRowsNormalized.length - 1;
        
        let actualDataRows = dataRowsNormalized;
        let truncated = false;
        
        if (dataEndRow > MAX_ROW) {
          // L200を超える分は切り捨て
          const maxDataRows = MAX_ROW - dataStartRow + 1;
          actualDataRows = dataRowsNormalized.slice(0, maxDataRows);
          truncated = true;
          warnings.push(`ファイル「${file.name}」: L200を超える分は切り捨てられました`);
        }
        
        // Excelに書き込み（範囲一括）
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getItem(month);
          
          // ファイル名行を書き込み
          const fileNameRange = sheet.getRange(`A${fileNameRowNum}:L${fileNameRowNum}`);
          fileNameRange.values = [fileNameRow];
          // ファイル名行の背景色を薄い青色に設定
          fileNameRange.format.fill.color = '#E6F3FF'; // 薄い青色
          
          // データ行を書き込み
          if (actualDataRows.length > 0) {
            const dataStartRowNum = fileNameRowNum + 1;
            const dataEndRowNum = dataStartRowNum + actualDataRows.length - 1;
            const dataRange = sheet.getRange(`A${dataStartRowNum}:L${dataEndRowNum}`);
            dataRange.values = actualDataRows;
          }
          
          await context.sync();
        });
        
        // 行ポインタを更新（最後に書いたデータ行の行番号）
        rowPtr = truncated ? MAX_ROW : (dataStartRow + actualDataRows.length - 1);
        sheetRowPtr[month] = rowPtr;
        
        console.log(`ファイル「${file.name}」: シート「${month}」に書き込み完了（行${fileNameRowNum}〜${rowPtr}）`);
      } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        warnings.push(`ファイル「${file.name}」: ${errorMessage}`);
        console.error(`ファイル「${file.name}」の処理でエラー:`, error);
      }
    }
    
    // 警告があれば表示
    if (warnings.length > 0) {
      const warningMessage = warnings.join('\n');
      // Officeアドイン環境では alert() が使えないため、エラーとして throw
      // App.tsx の handleFeatureClick でエラーメッセージとして表示される
      throw new Error(`以下の警告がありました:\n${warningMessage}`);
    }
    
    console.log('すべてのCSVインポートが完了しました');
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


