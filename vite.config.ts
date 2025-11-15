import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import fs from 'fs'
import os from 'os'
import path from 'path'

// Office Add-in開発用の証明書を読み込む
const certDir = path.join(os.homedir(), '.office-addin-dev-certs')
let httpsOptions: { cert: Buffer; key: Buffer } | undefined;

try {
  httpsOptions = {
    cert: fs.readFileSync(path.join(certDir, 'localhost.crt')),
    key: fs.readFileSync(path.join(certDir, 'localhost.key')),
  };
} catch (error) {
  console.warn('証明書ファイルが見つかりません。デフォルトのHTTPS設定を使用します。');
  console.warn('証明書を生成するには、以下のコマンドを実行してください:');
  console.warn('  npm run cert:install');
  // 証明書が見つからない場合は、デフォルトのHTTPS設定を使用
  httpsOptions = undefined;
}

// https://vite.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    https: (httpsOptions ?? true) as any,
    port: 5173,
  },
})
