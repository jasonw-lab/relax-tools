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
 */
export function registerFeature(feature: Feature): void {
  features.push(feature);
}

/**
 * カテゴリ別に機能を取得
 */
export function getFeaturesByCategory(): Record<FeatureCategory, Feature[]> {
  const result: Record<FeatureCategory, Feature[]> = {
    sheet: [],
    text: [],
    search: [],
    csv: [],
  };
  
  features.forEach(feature => {
    result[feature.category].push(feature);
  });
  
  return result;
}

/**
 * すべての機能を取得
 */
export function getAllFeatures(): Feature[] {
  return [...features];
}

