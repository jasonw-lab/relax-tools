/**
 * 検索・置換: 正規表現置換
 * アクティブシート全体に対して正規表現による置換を実行
 */

/* global Excel */

/**
 * アクティブシート全体に対して正規表現置換を実行
 * @param pattern 正規表現パターン
 * @param replacement 置換文字列（$1, $2などのグループ参照可）
 */
export async function regexReplaceAll(pattern: string, replacement: string): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      
      // 使用されている範囲を取得
      const usedRange = sheet.getUsedRange();
      usedRange.load(['address', 'values', 'rowCount', 'columnCount']);
      await context.sync();
      
      if (!usedRange.values || usedRange.rowCount === 0 || usedRange.columnCount === 0) {
        alert('データがありません');
        return;
      }
      
      const values = usedRange.values as (string | number | boolean | null)[][];
      const convertedValues: (string | number | boolean | null)[][] = [];
      
      // 正規表現を作成（グローバルフラグを追加）
      let regex: RegExp;
      try {
        // パターンがすでにフラグを含んでいる場合はそのまま、そうでなければ g フラグを追加
        regex = new RegExp(pattern, pattern.includes('/') ? '' : 'g');
      } catch (e) {
        alert('正規表現パターンが無効です: ' + pattern);
        throw e;
      }
      
      // 各行・各セルを処理
      for (const row of values) {
        const convertedRow: (string | number | boolean | null)[] = [];
        for (const cell of row) {
          if (typeof cell === 'string') {
            // 正規表現置換を実行
            const converted = cell.replace(regex, replacement);
            convertedRow.push(converted);
          } else {
            // 文字列以外はそのまま
            convertedRow.push(cell);
          }
        }
        convertedValues.push(convertedRow);
      }
      
      // 変換後の値を一括で書き込み
      usedRange.values = convertedValues;
      await context.sync();
      
      alert(`置換が完了しました（${usedRange.rowCount}行 x ${usedRange.columnCount}列）`);
    });
  } catch (error) {
    console.error('正規表現置換エラー:', error);
    alert('置換に失敗しました: ' + (error as Error).message);
    throw error;
  }
}

/**
 * ユーザー入力付きで正規表現置換を実行
 * プロンプトでパターンと置換文字列を入力させる
 */
export async function regexReplaceAllWithPrompt(): Promise<void> {
  const pattern = prompt('正規表現パターンを入力してください:');
  if (!pattern) {
    return;
  }
  
  const replacement = prompt('置換文字列を入力してください:') || '';
  
  await regexReplaceAll(pattern, replacement);
}

