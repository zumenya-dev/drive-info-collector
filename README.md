# Google Drive 共有ドライブ情報収集ツール

**バージョン:** 1.1
**作成日:** 2024年9月29日
**最終更新:** 2025年1月27日
**作成者:** Zumenya

## 📖 概要

Google Workspace環境の共有ドライブ情報を自動収集し、フォルダ構造、権限設定、外部共有状況を調査するツールです。

## 🎯 目的

- 全共有ドライブの棚卸しとリスト化
- ディレクトリ階層構造の完全な把握
- 権限設定と外部共有状況の調査
- 長期間未使用ドライブの特定
- セキュリティリスクの可視化

## ✨ 機能

### 📊 情報収集機能
- **共有ドライブ一覧**: 全ドライブの基本情報取得
- **統計情報**: ファイル数、フォルダ数、合計容量の自動集計
- **階層構造取得**: 最大10階層までの完全なフォルダツリー
- **ファイル情報**: 作成日、更新日、サイズ、作成者（lastModifyingUser/sharingUser対応）
- **権限分析**: Drive.Permissions.list による詳細権限情報取得（共有ドライブ完全対応）
- **外部共有検出**: 組織内ドメイン共有と真の外部共有を区別して表示

### 🛡️ セキュリティ機能
- **実行権限制御**: 指定ユーザーのみ実行可能
- **読み取り専用**: ファイル内容は一切変更しない
- **監査ログ**: 実行履歴とエラー記録
- **機密情報保護**: 個人情報やファイル内容は収集しない

### ⚡ パフォーマンス機能
- **バッチ処理**: API制限を考慮した効率的な処理
- **中断・再開**: 6分制限対応の分割実行
- **進捗表示**: リアルタイムの処理状況表示
- **エラー処理**: 堅牢なエラーハンドリング

## 📁 ファイル構成

```
Drive情報収集ツール/
├── コード.gs              # メインスクリプト
├── appsscript.json      # 設定ファイル
├── 実行手順書.md         # 詳細な実行手順
├── README.md           # このファイル

```

## 🚀 クイックスタート

### 1. 前提条件
- Google Workspace管理者権限
- Google Apps Scriptの利用権限
- 対象共有ドライブへのアクセス権

### 2. セットアップ（5分）
1. [script.google.com](https://script.google.com) で新規プロジェクト作成
2. `コード.gs` の内容をコピー&ペースト
3. `CONFIG` 設定を環境に合わせて変更（ALLOWED_USERS, COMPANY_DOMAIN）
4. Drive API (v3) を有効化
5. `appsscript.json` を設定して OAuth スコープを追加
6. 権限を承認して実行

### 3. 実行と結果確認
- 実行時間: 共有ドライブ数により変動（1ドライブあたり1-5分）
- 結果: 「Drive情報収集結果」スプレッドシート
- 進捗: リアルタイムで確認可能

## 📋 出力データ構造

### マスターシート: 共有ドライブ一覧
```
No | ドライブ名 | ID | 作成日 | ファイル数 | 容量(GB) | 外部共有              | 状況
1  | 営業部     | abc| 2023/1 | 1,234     | 5.2     | 組織内共有あり, 外部共有あり(3件) | 完了
2  | 技術部     | def| 2023/2 | 856       | 12.0    | 組織内共有あり        | 完了
3  | 管理部     | ghi| 2023/3 | 423       | 3.8     | なし                  | 完了
```

### 個別ドライブシート: 詳細情報
```
レベル | パス              | 種別   | 名前      | 作成者 | 権限（メールアドレス+ロール）                | 外部共有
0      | /                 | フォルダ| ルート    | 山田太郎 | org-admin@example.co.jp(organizer)        | 組織内共有あり
1      | /営業資料/         | フォルダ| 営業資料  | 田中次郎 | tanaka@example.co.jp(writer), team@example.co.jp(reader) | 組織内共有あり
2      | /営業資料/2024年/  | フォルダ| 2024年   | 田中次郎 | tanaka@example.co.jp(writer)              | なし
3      | /営業資料/.../提案書| ファイル| 提案書.xlsx| 佐藤花子 | sato@example.co.jp(owner), client@external.com(reader) | 外部共有あり(1件)
```

## 🔧 設定項目

### 基本設定
```javascript
const CONFIG = {
  SPREADSHEET_NAME: 'Drive情報収集結果',
  BATCH_SIZE: 1000,
  API_DELAY: 100,              // API呼び出し間の待機時間(ms)
  MAX_EXECUTION_TIME: 330,     // 最大実行時間(秒)
  MAX_DEPTH: 10,

  // 🔐 セキュリティ設定（変更必須）
  ALLOWED_USERS: [
    '*******@********.co.jp'   // 実行許可ユーザーのメールアドレス（複数可）
  ],

  // 🏢 会社設定（変更必須）
  COMPANY_DOMAIN: '********.co.jp'  // 外部共有判定に使用
};
```

**⚠️ 重要**: `COMPANY_DOMAIN` に設定したドメインは「組織内共有」として扱われ、外部共有としてカウントされません。
```

## 🛠️ 高度な使用方法

### カスタマイズオプション

#### 1. 収集対象の制限
```javascript
// 特定のドライブを除外
const excludeDrives = ['テスト用', 'アーカイブ'];

// 特定のファイル形式のみ収集
const targetMimeTypes = [
  'application/vnd.google-apps.document',
  'application/vnd.google-apps.spreadsheet'
];
```

#### 2. 出力フォーマットの変更
```javascript
// 日付フォーマット変更
const formatDate = (dateString) => {
  return new Date(dateString).toLocaleDateString('ja-JP');
};

// ファイルサイズ表示変更
const formatSize = (bytes, unit = 'MB') => {
  // カスタム表示ロジック
};
```

#### 3. 分析機能の追加
```javascript
// 重複ファイル検出
function findDuplicateFiles(items) {
  const nameMap = new Map();
  items.forEach(item => {
    if (nameMap.has(item.name)) {
      // 重複処理
    }
  });
}

// 古いファイル検出
function findOldFiles(items, monthsThreshold = 12) {
  const cutoffDate = new Date();
  cutoffDate.setMonth(cutoffDate.getMonth() - monthsThreshold);

  return items.filter(item =>
    new Date(item.modifiedTime) < cutoffDate
  );
}
```

## 📊 分析事例

### 問題検出のポイント

#### 🔍 セキュリティリスク
- **真の外部共有**（「外部共有あり」表示）ファイルの特定と権限レベル確認
- **組織内共有**（「組織内共有あり」表示）は通常問題なし
- 退職者アカウントの残存権限検出
- 過度に広い権限設定の発見
- 「全員」（type=anyone）への公開共有の検出

#### 📈 使用量分析
- 容量使用率上位ドライブの特定
- 長期間未使用ファイルの検出
- ファイル形式別の分布分析

#### 🏗️ 構造分析
- 深すぎる階層構造の検出
- 重複する可能性のあるフォルダ名
- 命名規則の一貫性チェック

## 🚨 制限事項と注意点

### API制限
- **1日のAPI呼び出し数**: Drive API の制限内（ファイルごとに Permissions.list を呼び出すため注意）
- **実行時間制限**: 6分（自動分割実行で対応、`continueExecution()` で再開可能）
- **同時実行制限**: 1プロジェクトあたり1実行
- **必要な OAuth スコープ**:
  - `https://www.googleapis.com/auth/drive`
  - `https://www.googleapis.com/auth/spreadsheets`
  - `https://www.googleapis.com/auth/userinfo.email`

### データ制限
- **最大階層深度**: 10階層（CONFIG.MAX_DEPTH で変更可能）
- **ファイル数制限**: 実質的な制限なし（処理時間に影響）
- **スプレッドシート行制限**: 500万行
- **権限情報取得**: ファイルごとに Drive.Permissions.list を呼び出すため、大量のファイルがある場合は処理時間が長くなる可能性あり

### セキュリティ注意事項
- 機密データを含む可能性があるため、結果の取り扱いに注意
- 実行ユーザーの制限を必ず設定
- 不要になった結果データは適切に削除

## 🔧 トラブルシューティング

### よくある問題

#### 権限エラー
```
原因: 共有ドライブへのアクセス権限不足
解決: Google Workspace管理者による権限付与
```

#### タイムアウトエラー
```
原因: 処理対象が多すぎる
解決: continueExecution()で継続実行
```

#### API制限エラー
```
原因: 1日のAPI呼び出し数上限に達した
解決: 翌日に再実行、またはAPI_DELAYを増やす
```

### パフォーマンス最適化

1. **BATCH_SIZE調整**: メモリ使用量と処理速度のバランス
2. **API_DELAY設定**: レート制限回避のための待機時間
3. **MAX_DEPTH制限**: 深い階層の処理時間短縮

## 📄 ライセンス

このツールはGIC社内での使用に限定されます。
第三者への提供や外部での使用は禁止されています。

## 📝 変更履歴

| バージョン | 日付 | 変更内容 |
|-----------|------|----------|
| 1.1 | 2025/01/27 | 権限取得機能強化（Drive.Permissions.list使用）、組織内/外部共有の区別表示、統計情報収集、OAuth スコープ最適化 |
| 1.0 | 2024/09/29 | 初回リリース |

---