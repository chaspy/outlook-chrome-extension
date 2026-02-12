# Outlook Chrome Extension

Outlook Web の予定表操作を補助する Chrome 拡張です。  
This is a Chrome extension that helps you work with Outlook Web Calendar.

## 日本語 (Japanese First)

### これは何？

この拡張は Outlook Web の予定表で、次の作業を楽にします。

- 重複予定の候補をハイライト表示
- 予定表リストを「名前 / メール / ID」で検索
- 連絡先リスト（CSV/TSV）を登録して候補表示に使う
- 予定作成時の出席者入力を補助
- トラブル調査用のデバッグ情報を取得

### まず必要なもの

- Google Chrome
- Outlook Web (`outlook.office.com`, `outlook.office365.com`, `outlook.live.com`) の利用権限

### 5分セットアップ（非エンジニア向け）

1. Releases ページを開く: `https://github.com/chaspy/outlook-chrome-extension/releases`
2. 最新リリースの `Assets` から `outlook-chrome-extension-vX.Y.Z.zip` をダウンロードする  
   (`Source code (zip)` ではなく、`outlook-chrome-extension-...zip` を選んでください)
3. ダウンロードした ZIP を展開する
4. Chrome で `chrome://extensions/` を開く
5. 右上の「デベロッパー モード」を ON
6. 「パッケージ化されていない拡張機能を読み込む」をクリック
7. 展開したフォルダ（`manifest.json` があるフォルダ）を選ぶ
8. Outlook Web の予定表ページを再読み込みする

### 使い方

#### 1. 重複予定をチェックする

1. Outlook Web の予定表を開く
2. 画面上に出る `重複検出` ボタンを押す
3. 重複候補がハイライトされる
4. 解除したいときは `重複をクリア` を押す

#### 2. 予定表リストを検索する

- 予定表リストの上に `検索（名前/メール/ID）` が表示されます
- 入力すると一致候補だけが強調されます
- 候補カードから `選択` / `解除` できます

#### 3. 連絡先CSV/TSVを登録する

1. Chrome の拡張アイコンから `Outlook Debug` を開く
2. `連絡先CSV/TSV` 欄に貼り付ける
3. `保存` を押す

入力形式（タブ / カンマ / セミコロン区切りに対応）:

```text
full_name,email_address,id
山田 太郎,taro.yamada@example.com,yamada
Suzuki Hanako,suzuki.hanako@example.com,suzuki
```

- ヘッダー行はあってもなくても動作します
- `name` と `email` が必須です
- `id` は任意です

#### 4. デバッグ情報を取得する

拡張ポップアップの `Debug` セクションで:

- `取得`: 現在のタブから状態を取得
- `コピー`: JSON をクリップボードへコピー

### うまく動かないとき

- ボタンや検索欄が出ない: Outlook の予定表ページを再読み込み
- `取得失敗` が出る: Outlook タブをアクティブにして再実行
- 拡張更新後に挙動が変: `chrome://extensions/` で再読み込み

### データについて

- 登録した連絡先は `chrome.storage.local` に保存されます
- 外部サーバー送信処理は実装していません（ブラウザ内保存のみ）

---

## English

### What is this?

This extension helps you on Outlook Web Calendar with:

- Highlighting possible overlapping events
- Searching calendar lists by name/email/ID
- Using saved contact CSV/TSV data for suggestions
- Assisting attendee input while creating events
- Collecting debug information when troubleshooting

### Requirements

- Google Chrome
- Access to Outlook Web (`outlook.office.com`, `outlook.office365.com`, `outlook.live.com`)

### 5-minute setup (for non-engineers)

1. Open Releases: `https://github.com/chaspy/outlook-chrome-extension/releases`
2. In the latest release `Assets`, download `outlook-chrome-extension-vX.Y.Z.zip`  
   (Do not choose `Source code (zip)`; choose `outlook-chrome-extension-...zip`.)
3. Extract the downloaded ZIP
4. Open `chrome://extensions/` in Chrome
5. Turn on `Developer mode`
6. Click `Load unpacked`
7. Select the extracted folder (the one that contains `manifest.json`)
8. Reload the Outlook Web Calendar page

### How to use

#### 1. Check overlapping events

1. Open Outlook Web Calendar
2. Click `重複検出` (Detect overlaps)
3. Possible overlaps are highlighted
4. Click `重複をクリア` to remove highlights

#### 2. Search calendar lists

- A search box (`検索（名前/メール/ID）`) appears above the calendar list
- Matching rows are highlighted as you type
- You can `選択` (select) / `解除` (deselect) directly from hint cards

#### 3. Save contacts from CSV/TSV

1. Open the extension popup (`Outlook Debug`)
2. Paste data into `連絡先CSV/TSV`
3. Click `保存` (Save)

Supported delimiters: tab, comma, semicolon.

```text
full_name,email_address,id
Taro Yamada,taro.yamada@example.com,yamada
Hanako Suzuki,suzuki.hanako@example.com,suzuki
```

- Header row is optional
- `name` and `email` are required
- `id` is optional

#### 4. Collect debug data

In the popup `Debug` section:

- `取得` (Fetch): collect status from the active tab
- `コピー` (Copy): copy JSON output to clipboard

### Troubleshooting

- No button/search box: reload Outlook Calendar page
- `取得失敗` (Fetch failed): activate an Outlook tab and retry
- Strange behavior after updates: reload extension in `chrome://extensions/`

### Data handling

- Saved contacts are stored in `chrome.storage.local`
- No external upload/sync logic is implemented

---

## For Developers

```bash
npm install
npm run lint
npm run lint:webext
npm run test:search
```
