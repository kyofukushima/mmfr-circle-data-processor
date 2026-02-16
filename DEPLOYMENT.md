# Streamlit Cloudデプロイ手順（チャット機能なし版）

このドキュメントでは、育児サークル情報処理アプリ（チャット機能なし版）をStreamlit Cloudにデプロイする手順を説明します。

**チャット機能なし版のため、OpenAI APIキーやsecrets.tomlの設定は不要です。**

## 前提条件

- GitHubアカウント
- Streamlit Cloudアカウント（GitHubアカウントでサインアップ可能）
- リポジトリがパブリックまたはStreamlit Cloudからアクセス可能

## 必要ファイル

以下のファイルがリポジトリに含まれている必要があります：

### 必須ファイル
- `app.py` - メインアプリケーション
- `requirements.txt` - Python依存関係
- `template.xlsx` - Excelテンプレートファイル
- `validate.py` - URL検証機能

### 設定ファイル
- `.streamlit/config.toml` - Streamlit設定
- `packages.txt` - システムレベル依存関係（空でも必要）

### 推奨ファイル
- `README.md` - プロジェクト説明
- `.gitignore` - Git除外設定
- `DEPLOYMENT.md` - このファイル

## デプロイ手順

### 1. GitHubリポジトリの準備

```bash
# リポジトリの初期化（まだの場合）
git init

# 全ファイルをステージング
git add .

# コミット
git commit -m "Initial commit for Streamlit Cloud deployment (no-chat version)"

# GitHubリポジトリにプッシュ
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
git branch -M main
git push -u origin main
```

### 2. Streamlit Cloudでのデプロイ

1. **Streamlit Cloud** ([share.streamlit.io](https://share.streamlit.io)) にアクセス
2. GitHubアカウントでサインイン
3. **"New app"** をクリック
4. **"From existing repo"** を選択
5. 以下の情報を入力：
   - **Repository**: `YOUR_USERNAME/YOUR_REPOSITORY`
   - **Branch**: `main`
   - **Main file path**: `app.py`
   - **App URL** (optional): カスタムURL

### 3. デプロイ設定の確認

- **Python version**: 3.9以上が推奨
- **Dependencies**: `requirements.txt`から自動インストール
- **System packages**: `packages.txt`から自動インストール
- **Secrets**: 不要（チャット機能なし版のため）

### 4. デプロイ完了

- デプロイプロセスが開始されます（通常2-5分）
- 完了すると、アプリのURLが提供されます
- エラーが発生した場合は、ログを確認して修正

## 設定詳細

### requirements.txt
```
streamlit>=1.34.0
pandas>=2.0.0
openpyxl>=3.1.0
chardet>=5.0.0
aiohttp>=3.8.0
requests>=2.28.0
```

※ openai、Pillowはチャット機能削除のため含まれていません。

### .streamlit/config.toml
```toml
[server]
maxUploadSize = 200
maxMessageSize = 200

[theme]
primaryColor = "#FF6B6B"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#F0F2F6"
textColor = "#262730"

[browser]
gatherUsageStats = false
```

## トラブルシューティング

### よくある問題

1. **ModuleNotFoundError**
   - `requirements.txt`に必要なパッケージが記載されているか確認
   - パッケージ名とバージョンが正しいか確認

2. **File not found: template.xlsx**
   - `template.xlsx`がリポジトリのルートにあるか確認
   - `.gitignore`で除外されていないか確認（!template.xlsxで含める）

3. **Memory/CPU limits**
   - ファイルサイズ制限: 200MB
   - 処理時間制限: 約1時間
   - 同時ユーザー制限: Community版では制限あり

4. **Import errors**
   - 相対importではなく絶対importを使用
   - `validate.py`がリポジトリに含まれているか確認

### デバッグ方法

1. **ローカルでの動作確認**
   ```bash
   streamlit run app.py
   ```

2. **Streamlit Cloudログの確認**
   - アプリ管理画面でログを確認
   - エラーメッセージから原因を特定

## セキュリティ考慮事項

- アップロードファイルのサイズ制限を適切に設定
- ユーザー入力の検証を徹底
- チャット機能なし版のため、APIキー等の機密情報管理は不要

## 更新手順

```bash
# 変更をコミット
git add .
git commit -m "Update: description of changes"

# GitHubにプッシュ
git push origin main
```

Streamlit Cloudは自動的に変更を検出し、再デプロイを開始します。

## サポート

- [Streamlit Documentation](https://docs.streamlit.io/)
- [Streamlit Cloud Documentation](https://docs.streamlit.io/streamlit-cloud)
- [Community Forum](https://discuss.streamlit.io/)
