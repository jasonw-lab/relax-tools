/**
 * 文字列変換: 全角→半角
 * 選択範囲の全角英数・カナを半角に変換
 */

/* global Excel */

/**
 * 選択範囲の全角文字を半角に変換
 * - 全角英数 → 半角英数
 * - 全角カナ → 半角カナ
 */
export async function toHankaku(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.load(['address', 'values']);
      await context.sync();
      
      const values = range.values as (string | number | boolean | null)[][];
      const convertedValues: (string | number | boolean | null)[][] = [];
      
      // 各行・各セルを処理
      for (const row of values) {
        const convertedRow: (string | number | boolean | null)[] = [];
        for (const cell of row) {
          if (typeof cell === 'string') {
            // 全角→半角変換
            let converted = cell;
            
            // 全角英数 → 半角英数 (0xFF00-0xFF5E → 0x0020-0x007E)
            converted = converted.replace(/[！-～]/g, (char) => {
              const code = char.charCodeAt(0);
              if (code >= 0xFF01 && code <= 0xFF5E) {
                return String.fromCharCode(code - 0xFEE0);
              }
              return char;
            });
            
            // 全角スペース → 半角スペース
            converted = converted.replace(/　/g, ' ');
            
            // 全角カナ → 半角カナ (カタカナ範囲: 0x30A0-0x30FF)
            // 簡易実装: 主要なカタカナのみ変換
            const kanaMap: Record<string, string> = {
              'ア': 'ｱ', 'イ': 'ｲ', 'ウ': 'ｳ', 'エ': 'ｴ', 'オ': 'ｵ',
              'カ': 'ｶ', 'キ': 'ｷ', 'ク': 'ｸ', 'ケ': 'ｹ', 'コ': 'ｺ',
              'サ': 'ｻ', 'シ': 'ｼ', 'ス': 'ｽ', 'セ': 'ｾ', 'ソ': 'ｿ',
              'タ': 'ﾀ', 'チ': 'ﾁ', 'ツ': 'ﾂ', 'テ': 'ﾃ', 'ト': 'ﾄ',
              'ナ': 'ﾅ', 'ニ': 'ﾆ', 'ヌ': 'ﾇ', 'ネ': 'ﾈ', 'ノ': 'ﾉ',
              'ハ': 'ﾊ', 'ヒ': 'ﾋ', 'フ': 'ﾌ', 'ヘ': 'ﾍ', 'ホ': 'ﾎ',
              'マ': 'ﾏ', 'ミ': 'ﾐ', 'ム': 'ﾑ', 'メ': 'ﾒ', 'モ': 'ﾓ',
              'ヤ': 'ﾔ', 'ユ': 'ﾕ', 'ヨ': 'ﾖ',
              'ラ': 'ﾗ', 'リ': 'ﾘ', 'ル': 'ﾙ', 'レ': 'ﾚ', 'ロ': 'ﾛ',
              'ワ': 'ﾜ', 'ヲ': 'ｦ', 'ン': 'ﾝ',
              'ァ': 'ｧ', 'ィ': 'ｨ', 'ゥ': 'ｩ', 'ェ': 'ｪ', 'ォ': 'ｫ',
              'ッ': 'ｯ', 'ャ': 'ｬ', 'ュ': 'ｭ', 'ョ': 'ｮ',
              'ガ': 'ｶﾞ', 'ギ': 'ｷﾞ', 'グ': 'ｸﾞ', 'ゲ': 'ｹﾞ', 'ゴ': 'ｺﾞ',
              'ザ': 'ｻﾞ', 'ジ': 'ｼﾞ', 'ズ': 'ｽﾞ', 'ゼ': 'ｾﾞ', 'ゾ': 'ｿﾞ',
              'ダ': 'ﾀﾞ', 'ヂ': 'ﾁﾞ', 'ヅ': 'ﾂﾞ', 'デ': 'ﾃﾞ', 'ド': 'ﾄﾞ',
              'バ': 'ﾊﾞ', 'ビ': 'ﾋﾞ', 'ブ': 'ﾌﾞ', 'ベ': 'ﾍﾞ', 'ボ': 'ﾎﾞ',
              'パ': 'ﾊﾟ', 'ピ': 'ﾋﾟ', 'プ': 'ﾌﾟ', 'ペ': 'ﾍﾟ', 'ポ': 'ﾎﾟ',
              'ヴ': 'ｳﾞ',
            };
            
            // カタカナ変換（簡易版）
            for (const [zenkaku, hankaku] of Object.entries(kanaMap)) {
              converted = converted.replace(new RegExp(zenkaku, 'g'), hankaku);
            }
            
            convertedRow.push(converted);
          } else {
            // 文字列以外はそのまま
            convertedRow.push(cell);
          }
        }
        convertedValues.push(convertedRow);
      }
      
      // 変換後の値を一括で書き込み（セル単位ループを避ける）
      range.values = convertedValues;
      await context.sync();
    });
  } catch (error) {
    console.error('全角→半角変換エラー:', error);
    alert('変換に失敗しました: ' + (error as Error).message);
    throw error;
  }
}

