/**
 * シート操作: 消費種類自動設定
 * アクティブシートの4行目から最終行まで、B列（利用店名/商品名）をチェックして
 * K列（消費種類）に自動設定する
 */

/* global Excel */

/**
 * type.csvを読み込んでconsumptionMapを構築
 * @returns keyword -> type のマッピング（Map）
 */
async function loadConsumptionMap(): Promise<Map<string, string>> {
  try {
    const response = await fetch('/type.csv');
    if (!response.ok) {
      throw new Error('type.csvの読み込みに失敗しました');
    }
    const text = await response.text();
    
    // Papa.parse でCSVをパース
    const Papa = (await import('papaparse')).default;
    const parseResult = Papa.parse(text, {
      header: true,
      skipEmptyLines: 'greedy' as any,
    }) as { data: Array<{ keyword: string; type: string }>; errors: any[] };
    
    if (parseResult.errors.length > 0) {
      console.warn('type.csvパース警告:', parseResult.errors);
    }
    
    const consumptionMap = new Map<string, string>();
    for (const row of parseResult.data) {
      // 第1列（keyword）と第2列（type）を使用
      if (row.keyword && row.type) {
        consumptionMap.set(row.keyword.trim(), row.type.trim());
      }
    }
    
    console.log(`consumptionMapを構築しました: ${consumptionMap.size}件のマッピング`);
    return consumptionMap;
  } catch (error) {
    console.error('type.csvの読み込みでエラーが発生しました:', error);
    throw new Error(`type.csvの読み込みに失敗しました: ${error instanceof Error ? error.message : String(error)}`);
  }
}

/**
 * 消費種類自動設定機能
 * アクティブシートの4行目から最終行まで、B列をチェックしてK列に書き込む
 */
export async function setConsumptionCategory(): Promise<void> {
  try {
    console.log('消費種類自動設定開始');
    
    // type.csvからconsumptionMapを構築
    const consumptionMap = await loadConsumptionMap();
    
    if (consumptionMap.size === 0) {
      throw new Error('type.csvに有効なマッピングがありません');
    }
    
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // 使用範囲を取得して最終行を特定
      const usedRange = sheet.getUsedRange();
      usedRange.load(['rowCount', 'rowIndex']);
      await context.sync();
      
      const startRow = 4; // 4行目から開始（1-based）
      const endRow = usedRange.rowIndex + usedRange.rowCount - 1; // 最終行（1-based）
      
      if (endRow < startRow) {
        console.log('処理対象の行がありません');
        return;
      }
      
      console.log(`処理範囲: ${startRow}行目〜${endRow}行目`);
      
      // B列とK列のデータを一括取得（4行目から最終行まで）
      const bColumnRange = sheet.getRange(`B${startRow}:B${endRow}`);
      const kColumnRange = sheet.getRange(`K${startRow}:K${endRow}`);
      
      bColumnRange.load('values');
      kColumnRange.load('values');
      await context.sync();
      
      const bValues = bColumnRange.values as (string | number | boolean | null)[][];
      const kValues = kColumnRange.values as (string | number | boolean | null)[][];
      
      // 新しいK列の値を構築
      const newKValues: (string | number | boolean | null)[][] = [];
      let matchedCount = 0;
      
      for (let i = 0; i < bValues.length; i++) {
        const bValue = String(bValues[i][0] || '').trim();
        const currentKValue = kValues[i][0];
        
        // B列が空の場合はスキップ（K列は変更しない）
        if (!bValue) {
          newKValues.push([currentKValue]);
          continue;
        }
        
        // consumptionMapの各キー（キーワード）がB列文字列に部分一致するかチェック
        let matchedType: string | null = null;
        for (const [keyword, type] of consumptionMap.entries()) {
          if (bValue.includes(keyword)) {
            matchedType = type;
            matchedCount++;
            break; // 最初にマッチしたキーを使用
          }
        }
        
        // マッチした場合は新しい値を設定、マッチしない場合は既存の値を維持
        if (matchedType !== null) {
          newKValues.push([matchedType]);
        } else {
          newKValues.push([currentKValue]);
        }
      }
      
      // K列に一括書き込み
      if (matchedCount > 0) {
        kColumnRange.values = newKValues;
        await context.sync();
        console.log(`消費種類自動設定完了: ${matchedCount}件のマッチ`);
      } else {
        console.log('マッチするキーワードが見つかりませんでした');
      }
    });
  } catch (error) {
    console.error('消費種類自動設定エラー:', error);
    const errorMessage = error instanceof Error ? error.message : String(error);
    throw new Error(`消費種類の自動設定に失敗しました: ${errorMessage}`);
  }
}

