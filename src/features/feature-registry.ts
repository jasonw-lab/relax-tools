/**
 * 機能レジストリ
 * カテゴリ別に機能を管理し、UIに表示するための定義
 */

export type FeatureCategory = 'sheet' | 'text' | 'search' | 'csv';

export interface Feature {
  id: string;
  name: string;
  category: FeatureCategory;
  description: string;
  execute: () => Promise<void>;
}

// 機能のカテゴリ名表示用
export const categoryLabels: Record<FeatureCategory, string> = {
  sheet: 'シート操作',
  text: '文字列変換',
  search: '検索・置換',
  csv: 'CSV操作',
};

// 機能レジストリ（グローバル配列）
const features: Feature[] = [];

/**
 * 機能をレジストリに登録
 * 同じIDの機能が既に登録されている場合は登録をスキップ
 */
export function registerFeature(feature: Feature): void {
  // 同じIDの機能が既に登録されているかチェック
  const existingIndex = features.findIndex(f => f.id === feature.id);
  if (existingIndex >= 0) {
    // 既に登録されている場合は置き換え
    features[existingIndex] = feature;
  } else {
    // 新規登録
    features.push(feature);
  }
}

/**
 * カテゴリ別に機能を取得
 * 重複を排除して返す
 */
export function getFeaturesByCategory(): Record<FeatureCategory, Feature[]> {
  const result: Record<FeatureCategory, Feature[]> = {
    sheet: [],
    text: [],
    search: [],
    csv: [],
  };
  
  // 重複を排除するために、IDでマップを作成
  const seenIds = new Set<string>();
  
  features.forEach(feature => {
    // 同じIDの機能が既に追加されている場合はスキップ
    if (!seenIds.has(feature.id)) {
      seenIds.add(feature.id);
      result[feature.category].push(feature);
    }
  });
  
  return result;
}

/**
 * すべての機能を取得
 */
export function getAllFeatures(): Feature[] {
  return [...features];
}

