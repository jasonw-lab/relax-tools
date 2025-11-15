# Excel Tools (Office Add-in)

Excel用のOffice Add-inです。React + TypeScript + Viteで構築されています。

## 必要な環境

- **Node.js**: 20.19以上 または 22.12以上（Vite 7.x の要件）
  - 現在のバージョンを確認: `node --version`
  - アップグレードが必要な場合: [Node.js公式サイト](https://nodejs.org/)から最新版をインストール

## セットアップ

### 1. 依存関係のインストール

```bash
npm install
```

### 2. 開発用証明書の生成（初回のみ）

Office Add-inの開発にはHTTPS証明書が必要です。以下のコマンドで証明書を生成します：

```bash
npm run cert:install
```

証明書の生成後、macOSの場合はキーチェーンへの信頼設定が必要な場合があります。

## 起動・実行・デバッグ方法

### 開発サーバーの起動

```bash
npm run dev
```

### デバッグ実行

開発サーバーを起動した後、**別のターミナルタブ**で以下を実行：

```bash
npm run debug:desktop
```

これにより、Excel Desktopアプリケーションでアドインが起動し、デバッグが可能になります。

### デバッグの停止

```bash
npm run stop
```
