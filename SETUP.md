# Calendar Availability Finder v2 - セットアップ手順

## アーキテクチャ

Chrome拡張機能 → **Webアプリ（静的HTML/JS/CSS）** に変更。
サーバーにデプロイすれば、ローカルファイル削除の影響を受けずにブラウザからアクセスできます。

## ファイル構成

```
index.html   ... メインページ
app.js       ... アプリロジック（OAuth2, Calendar API, 空き時間計算）
app.css      ... スタイル
mock.html    ... デモ用モック表示
SETUP.md     ... このファイル
```

## 1. Google Cloud Console でプロジェクト設定

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. プロジェクトを作成（または既存を使用）
3. 「APIとサービス」→「ライブラリ」→ **Google Calendar API** を有効化

## 2. OAuth 2.0 クライアントIDを作成

1. 「APIとサービス」→「認証情報」→「認証情報を作成」→「OAuthクライアントID」
2. アプリケーションの種類: **ウェブアプリケーション**
3. 名前: `Calendar Availability Finder`
4. **承認済みの JavaScript 生成元** にデプロイ先URLを追加:
   - ローカルテスト: `http://localhost:8080`
   - 本番: `https://your-domain.com`
5. クライアントIDをコピー

## 3. app.js の設定を更新

`app.js` 冒頭の `CONFIG.CLIENT_ID` を書き換え:

```js
const CONFIG = {
  CLIENT_ID: 'あなたのクライアントID.apps.googleusercontent.com',
  ...
};
```

## 4. OAuth同意画面の設定

1. 「APIとサービス」→「OAuth同意画面」
2. ユーザータイプ: 外部（組織内なら内部も可）
3. スコープ追加:
   - `https://www.googleapis.com/auth/calendar.readonly`
   - `https://www.googleapis.com/auth/calendar.freebusy`
4. テストユーザーに利用者のGoogleアカウントを追加

## 5. デプロイ方法

### 方法A: GitHub Pages（無料・簡単）

1. GitHubにリポジトリを作成
2. `index.html`, `app.js`, `app.css` をプッシュ
3. Settings → Pages → Source を `main` ブランチに設定
4. `https://username.github.io/repo-name/` でアクセス可能
5. このURLをOAuth設定の「承認済みのJavaScript生成元」に追加

### 方法B: Vercel / Netlify（無料）

1. GitHubリポジトリを連携するだけで自動デプロイ
2. 発行されたURLをOAuth設定に追加

### 方法C: 社内サーバー

1. Webサーバー（nginx, Apache等）に3ファイルを配置
2. HTTPS必須（Google OAuthの要件）
3. URLをOAuth設定に追加

### ローカルテスト

```bash
# Python 3
python -m http.server 8080

# Node.js
npx serve -p 8080
```

`http://localhost:8080` でアクセス

## 使い方

1. ブラウザでアプリURLにアクセス
2. 「Googleアカウントでログイン」をクリック
3. 参加者のメールアドレスを入力して追加
4. 設定（検索範囲、時間帯、除外キーワード等）をカスタマイズ
5. 「空き時間を検索」をクリック
6. 結果が「競合なし」「競合あり（少ない順）」に分かれて表示

## v2 改善点

- **① 除外キーワード**: 予定名に「画面操作」等を含む場合は空き扱い（設定でカスタム可能）
- **② 競合表示の限定**: 検索時間帯（例: 11:00-18:00の平日）内の競合のみ表示
- **③ 2段階の候補表示**: 「競合なし」と「競合が少ない順」の2セクションに分割
- **④ Webアプリ化**: サーバーにデプロイ可能。ローカルファイル削除の影響なし

## 注意事項

- FreeBusy APIは、対象のGoogleカレンダーが「空き時間情報を共有」に設定されている必要があります
- 組織外のユーザーのカレンダーにアクセスする場合、対象ユーザーのカレンダー共有設定に依存します
- テスト段階ではOAuth同意画面で「テストユーザー」に追加したアカウントのみ認証可能です
