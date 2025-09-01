# ScrapeGAS

Google Apps Script（GAS）を使って、Webページを解析してスプレッドシートの表に出力します

---

Node.js v20をインストールするには、以下の方法が最も簡単で確実です。OSごとに手順を分けて紹介します。


## 🪟 Windowsの場合

### 方法①：公式インストーラーを使う（推奨）

1. 👉 [Node.js公式ダウンロードページ](https://nodejs.org/en/download)にアクセス
2. 「Current」タブから **Node.js v20.x.x** を選択
3. `.msi` インストーラーをダウンロード
4. ダブルクリックしてインストール（「Next」を押して進めるだけ）

### 方法②：ZIP版を使う（環境変数を手動設定）

1. [こちらの解説](https://www.flavor-of-life.com/entry/how_to_nodejs18_20_zip_install)によると：
   - ZIPファイルを任意のフォルダに解凍（例：`C:\node\node-v20.x.x-win-x64`）
   - システム環境変数 `Path` にそのフォルダを追加
   - コマンドプロンプトで `node -v` を実行して確認

---

## 🐧 macOS / Linuxの場合

### 方法①：`nvm`（Node Version Manager）を使う

```bash
curl -o- https://raw.githubusercontent.com/nvm-sh/nvm/v0.39.5/install.sh | bash
source ~/.bashrc  # または ~/.zshrc
nvm install 20
nvm use 20
```

これで `node -v` でバージョン確認できます。

### 方法②：公式サイトからバイナリをダウンロード

- [Node.js公式サイト](https://nodejs.org/en/download)から `.tar.gz` を取得
- 解凍して `/usr/local` などに配置
- `PATH` を通す必要があります

---

### 方法③：[nodesource.comが提供しているNode.jsのPPAを追加してインストールする。 ](https://redj.hatenablog.com/entry/2024/02/20/011007)

```bash
curl -fsSL https://deb.nodesource.com/setup_20.x | sudo -E bash - && sudo apt-get install -y nodejs
```

## ✅ インストール確認

```bash
node -v   # → v20.x.x
npm -v    # → npmのバージョンも確認
```

---


### 🥈 npmのグローバルディレクトリをユーザー領域に変更（安全で推奨）

```bash
mkdir ~/.npm-global
npm config set prefix '~/.npm-global'
```

次に、環境変数 `PATH` を更新：

#### macOS/Linux の場合（bash/zsh）

```bash
echo 'export PATH=$HOME/.npm-global/bin:$PATH' >> ~/.bashrc
source ~/.bashrc
```

または `~/.zshrc` に追加して `source ~/.zshrc`

#### Windows の場合（PowerShell）

```powershell
$env:PATH += ";$HOME\.npm-global\bin"
```

その後、インストール：

```bash
npm install -g @google/clasp
```

---


これで、`sudo` なしで安全にインストールできます。

---



## ✅ ユーザー設定でApps Script APIを有効化

### 🔧 ステップ①：設定ページにアクセス

👉 [https://script.google.com/home/usersettings](https://script.google.com/home/usersettings)

### 🟢 ステップ②：「Apps Script API」スイッチをオンにする

- ページ内に「Apps Script API」という項目があります
- スイッチを **ON（有効）** にしてください

> ⚠️ Google Workspace（旧G Suite）アカウントの場合、管理者がこの設定を制限していることがあります。その場合は管理者に問い合わせが必要です。

---

### ⏳ ステップ③：数分待ってから再試行

APIの有効化が反映されるまで、**数分〜10分程度**かかることがあります。

その後、以下のコマンドを実行：

```bash
clasp create --title "ScrapeGAS" --type sheets
```

---

## 🧭 補足：確認ポイント

| 項目 | 内容 |
|------|------|
| アカウント | `clasp login` で使用しているGoogleアカウントと一致しているか |
| API有効化 | ユーザー設定ページでApps Script APIがONになっているか |
| 反映待ち | 有効化後すぐに試すと反映されていない可能性あり |

---

## 🎯 うまくいったら次にできること

- `clasp push` でローカルコードをアップロード
- `clasp pull` で既存プロジェクトを取得
- Gitでバージョン管理
- VS Codeなどで快適にGAS開発

---

Google Apps Script

<img width="817" height="254" alt="image" src="https://github.com/user-attachments/assets/de2f2962-9683-40d4-85e2-826181f2909b" />

<img width="840" height="680" alt="image" src="https://github.com/user-attachments/assets/c105c410-6d2d-4852-ba1f-70b246e289bd" />

---

scriptId

<img width="655" height="111" alt="image" src="https://github.com/user-attachments/assets/4aad9c4c-fe22-4dd1-89fc-bdf5559d84d2" />

---

clasp push

<img width="556" height="89" alt="image" src="https://github.com/user-attachments/assets/a7611745-eec3-4299-ba86-5711cb1cdd7a" />

---

clasp deploy

<img width="919" height="128" alt="image" src="https://github.com/user-attachments/assets/0e5b47ec-96e9-46e4-a265-e59d9b20a665" />

---

<img width="596" height="227" alt="image" src="https://github.com/user-attachments/assets/c0ced4de-7fd0-4e8a-937d-151cf3badd86" />

<img width="652" height="790" alt="2025-07-15a" src="https://github.com/user-attachments/assets/3a8e945c-36a9-449a-a1bd-d2ac3c80dfb8" />

<img width="596" height="237" alt="image" src="https://github.com/user-attachments/assets/d1adc949-5fc7-434c-957e-b170f706f7f4" />

---
