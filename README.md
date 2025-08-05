# Yuine AI - 集合知AI システム

会議やセミナーでの参加者の意見を収集・分析し、集合知を生成するGoogle Apps Script (GAS) ベースのシステムです。

## 概要

Yuine AIは、会議やセミナーの参加者がAIと対話しながら意見を提供し、それらを集約・分析して集合知を生成するシステムです。リアルタイムでの意見収集と分析により、会議の効率化と参加者全体の知見共有を実現します。

## 主な機能

### 管理者機能
- **セッション作成・管理**: テーマ設定、質問生成、セッション開始・終了
- **AI質問生成**: Gemini AIを使用した自動質問生成
- **リアルタイム監視**: 参加者の回答状況をリアルタイムで監視
- **結果分析**: 参加者の回答を分析し、合意点・分散点を抽出

### 参加者機能
- **セッション参加**: QRコードまたはURLでの簡単参加
- **AI対話**: 段階的な質問への回答とAIフィードバック
- **プログレス表示**: 回答進捗の視覚的表示

### システム機能
- **データ管理**: Google Spreadsheetsでのデータ永続化
- **AI分析**: Gemini AIによる意見の分析・要約
- **エクスポート**: CSV形式での結果エクスポート

## 技術スタック

- **バックエンド**: Google Apps Script (GAS)
- **フロントエンド**: HTML/CSS/JavaScript
- **データベース**: Google Spreadsheets
- **AI**: Google Gemini API
- **デプロイ**: GAS WebApp

## セットアップ手順

### 1. Google Apps Scriptプロジェクト作成
1. [Google Apps Script](https://script.google.com/) にアクセス
2. 新しいプロジェクトを作成
3. `Code.gs`の内容をコピー＆ペースト

### 2. HTMLファイルの追加
以下のHTMLファイルをGASプロジェクトに追加：
- `admin.html` - 管理者画面
- `session.html` - 参加者セッション画面
- `monitor.html` - セッション監視画面
- `results.html` - 結果表示画面
- `templates.html` - テンプレート管理画面

### 3. Google Spreadsheet作成
1. 新しいGoogle Spreadsheetを作成
2. スプレッドシートIDを`Code.gs`の`SPREADSHEET_ID`に設定
3. 以下のシートを作成：
   - `Sessions` - セッション情報
   - `Participants` - 参加者情報
   - `Results` - 回答結果

### 4. Gemini API設定
1. [Google AI Studio](https://aistudio.google.com/) でAPIキーを取得
2. `Code.gs`の`GEMINI_API_KEY`に設定

### 5. デプロイ
1. GASプロジェクトで「デプロイ」→「新しいデプロイ」
2. 種類：「ウェブアプリ」
3. 実行者：「自分」
4. アクセス権：「全員」
5. デプロイIDを取得

## 使用方法

### 管理者操作
1. **セッション作成**
   - 管理者画面でテーマを入力
   - AI質問生成またはカスタム質問作成
   - セッション開始

2. **監視・分析**
   - 監視画面でリアルタイムでの参加状況確認
   - 結果画面で回答の分析・エクスポート

### 参加者操作
1. **セッション参加**
   - QRコードまたはURLでアクセス
   - セッションに参加

2. **質問回答**
   - 3段階の質問に順次回答
   - AIからのフィードバックを確認

## ファイル構成

```
yuine_ai/
├── Code.gs              # メインのGASコード
├── setup-config.gs     # セットアップ用設定ファイル
├── admin.html          # 管理者画面
├── session.html         # 参加者セッション画面
├── monitor.html        # セッション監視画面
├── results.html        # 結果表示画面
├── templates.html      # テンプレート管理画面
└── README.md           # このファイル
```

## 注意事項

- Google Apps ScriptとGemini APIの利用制限にご注意ください
- 大規模な利用の場合は、適切なエラーハンドリングとレート制限対策が必要です
- 機密情報を扱う場合は、適切なアクセス制御を設定してください

## ライセンス

MIT License

## 着想について

このアプリは深津さんのポッドキャストから着想を得ました。

**参考ポッドキャスト**: [深津さんのポッドキャスト](https://open.spotify.com/episode/0QJeULQ5hrRly9UWhxYQto?si=v_kI4PkuTUaLM_QxbTsMBg)

## 開発・メンテナンス

このプロジェクトは集合知の活用を目的としたオープンソースプロジェクトです。
バグ報告や機能要望は、GitHubのIssuesにてお願いします。

---

**🤖 Generated with [Claude Code](https://claude.ai/code)**