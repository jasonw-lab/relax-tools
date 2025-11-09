/**
 * メインアプリケーション
 * Office.js の初期化とタスクペインUI
 */

import { useEffect, useState, useRef } from 'react';
import './App.css';
import { 
  registerFeature, 
  getFeaturesByCategory, 
  categoryLabels, 
  type FeatureCategory 
} from './features/feature-registry';
import { removeBlankRowsInSelection } from './features/sheet/blank-ops';
import { toHankaku } from './features/text/convert';
import { regexReplaceAllWithPrompt } from './features/search/regex';
import { importCsvQuick } from './features/csv/importer';

// Office.js はグローバルに定義されている（index.html で読み込み）
// 型定義は @types/office-js から自動的に読み込まれる

function App() {
  const [isOfficeReady, setIsOfficeReady] = useState(false);
  const [message, setMessage] = useState<{ type: 'success' | 'error'; text: string } | null>(null);
  const [features, setFeatures] = useState<Record<FeatureCategory, ReturnType<typeof getFeaturesByCategory>[FeatureCategory]>>({
    sheet: [],
    text: [],
    search: [],
    csv: [],
  });
  const featuresRegisteredRef = useRef(false);

  useEffect(() => {
    // Office.js の初期化
    if (typeof Office !== 'undefined') {
      Office.onReady((info: any) => {
        if (info.host === Office.HostType.Excel) {
          setIsOfficeReady(true);
          
          // リボンコマンドの関連付け（Shared Runtime用）
          if (Office.context.requirements.isSetSupported('SharedRuntime')) {
            Office.actions.associate('showTaskPane', () => {
              // タスクペインは既に表示されているので、ここでは特に処理なし
              console.log('タスクペイン表示');
            });
          }
          
          // URLパラメータからアクションを確認（リボンから開かれた場合）
          const urlParams = new URLSearchParams(window.location.search);
          const action = urlParams.get('action');
          if (action === 'importCsv') {
            // CSVインポートを自動実行
            console.log('リボンからCSVインポートを実行');
            setTimeout(() => {
              importCsvQuick().catch((error) => {
                console.error('CSVインポートエラー:', error);
                setMessage({ type: 'error', text: `エラー: ${error instanceof Error ? error.message : String(error)}` });
                setTimeout(() => setMessage(null), 5000);
              });
            }, 500);
          }
          
          // 機能をレジストリに登録（一度だけ実行）
          if (!featuresRegisteredRef.current) {
            registerFeature({
              id: 'remove-blank-rows',
              name: '空行削除',
              category: 'sheet',
              description: '選択範囲内の空行を削除します',
              execute: removeBlankRowsInSelection,
            });
            
            registerFeature({
              id: 'to-hankaku',
              name: '全角→半角',
              category: 'text',
              description: '選択範囲の全角文字を半角に変換します',
              execute: toHankaku,
            });
            
            registerFeature({
              id: 'regex-replace',
              name: '正規表現置換',
              category: 'search',
              description: 'アクティブシート全体に対して正規表現置換を実行します',
              execute: regexReplaceAllWithPrompt,
            });
            
            registerFeature({
              id: 'csv-import',
              name: 'CSV取込（カード）',
              category: 'csv',
              description: 'CSVファイルを選択し、ファイル名から月を抽出して該当シートにインポートします',
              execute: importCsvQuick,
            });
            
            featuresRegisteredRef.current = true;
          }
          
          // カテゴリ別に機能を取得
          setFeatures(getFeaturesByCategory());
        }
      });
    }
  }, []);

  const handleFeatureClick = async (feature: ReturnType<typeof getFeaturesByCategory>[FeatureCategory][number]) => {
    try {
      console.log(`機能を実行中: ${feature.name}`);
      setMessage(null); // メッセージをクリア
      await feature.execute();
      console.log(`機能の実行が完了: ${feature.name}`);
      
      // 成功メッセージを設定（CSV取込など、特定の機能の場合）
      if (feature.id === 'csv-import') {
        setMessage({ type: 'success', text: 'CSVのインポートが完了しました' });
        // 3秒後に自動で消す
        setTimeout(() => setMessage(null), 3000);
      }
    } catch (error) {
      console.error(`機能実行エラー (${feature.name}):`, error);
      const errorMessage = error instanceof Error ? error.message : String(error);
      setMessage({ type: 'error', text: `エラー: ${errorMessage}` });
      // エラーメッセージは5秒後に消す
      setTimeout(() => setMessage(null), 5000);
    }
  };

  if (!isOfficeReady) {
    return (
      <div style={{ padding: '20px', textAlign: 'center' }}>
        <p>Excel を初期化中...</p>
      </div>
    );
  }

  return (
    <div style={{ padding: '16px' }}>
      <h1 style={{ fontSize: '18px', marginBottom: '20px' }}>Excel Tools</h1>
      
      {/* メッセージ表示 */}
      {message && (
        <div
          style={{
            padding: '12px',
            marginBottom: '16px',
            borderRadius: '4px',
            backgroundColor: message.type === 'success' ? '#d4edda' : '#f8d7da',
            color: message.type === 'success' ? '#155724' : '#721c24',
            border: `1px solid ${message.type === 'success' ? '#c3e6cb' : '#f5c6cb'}`,
            fontSize: '14px',
          }}
        >
          {message.text}
        </div>
      )}
      
      {/* カテゴリ別に機能を表示 */}
      {Object.entries(categoryLabels).map(([category, label]) => {
        const categoryFeatures = features[category as FeatureCategory];
        if (categoryFeatures.length === 0) {
          return null;
        }
        
        return (
          <div key={category} style={{ marginBottom: '24px' }}>
            <h2 style={{ fontSize: '14px', fontWeight: 'bold', marginBottom: '8px', color: '#333' }}>
              {label}
            </h2>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(2, 1fr)', gap: '8px' }}>
              {categoryFeatures.map((feature) => (
                <button
                  key={feature.id}
                  onClick={() => handleFeatureClick(feature)}
                  style={{
                    padding: '12px',
                    border: '1px solid #ddd',
                    borderRadius: '4px',
                    backgroundColor: '#fff',
                    cursor: 'pointer',
                    textAlign: 'left',
                    fontSize: '12px',
                  }}
                  onMouseEnter={(e) => {
                    e.currentTarget.style.backgroundColor = '#f0f0f0';
                  }}
                  onMouseLeave={(e) => {
                    e.currentTarget.style.backgroundColor = '#fff';
                  }}
                >
                  <div style={{ fontWeight: 'bold', marginBottom: '4px' }}>
                    {feature.name}
                  </div>
                  <div style={{ fontSize: '10px', color: '#666' }}>
                    {feature.description}
                  </div>
                </button>
              ))}
            </div>
          </div>
        );
      })}
    </div>
  );
}

export default App;
