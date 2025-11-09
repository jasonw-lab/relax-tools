/**
 * シート操作: 空行削除
 * 選択範囲内の空行を削除する
 */

/* global Excel */

/**
 * 選択範囲内の空行を削除
 * 空行 = すべてのセルが空文字列またはnullの行
 */
export async function removeBlankRowsInSelection(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'rowCount', 'columnCount', 'rowIndex', 'columnIndex']);
      await context.sync();
      
      // 選択範囲を取得
      const startRow = range.rowIndex;
      const rowCount = range.rowCount;
      const startCol = range.columnIndex;
      const colCount = range.columnCount;
      
      // 選択範囲のデータを一度に取得（範囲全体を読み込む方が効率的）
      const dataRange = range.worksheet.getRange(`${range.address}`);
      dataRange.load('values');
      await context.sync();
      
      const values = dataRange.values as (string | number | boolean | null)[][];
      
      // 空行のインデックスを収集（範囲内での相対インデックス）
      const blankRowIndices: number[] = [];
      for (let i = 0; i < values.length; i++) {
        const row = values[i];
        // 行内のすべてのセルが空の場合は空行
        const isEmpty = row.every(cell => cell === null || cell === '' || cell === undefined);
        if (isEmpty) {
          blankRowIndices.push(i);
        }
      }
      
      if (blankRowIndices.length === 0) {
        // 空行がなければ何もしない
        return;
      }
      
      // 空行を削除（下から上へ削除することでインデックスのズレを防ぐ）
      const sheet = range.worksheet;
      // シート全体の列数を取得（選択範囲の列数を使用）
      const maxCol = startCol + colCount - 1;
      
      // 逆順で処理（下から上へ）
      for (let i = blankRowIndices.length - 1; i >= 0; i--) {
        const relativeRowIndex = blankRowIndices[i];
        const actualRowIndex = startRow + relativeRowIndex;
        // 行全体を削除（選択範囲の列のみ）
        const rowRange = sheet.getRangeByIndexes(actualRowIndex, startCol, 1, colCount);
        rowRange.delete(Excel.DeleteShiftDirection.up);
      }
      
      await context.sync();
    });
  } catch (error) {
    console.error('空行削除エラー:', error);
    alert('空行の削除に失敗しました: ' + (error as Error).message);
    throw error;
  }
}

