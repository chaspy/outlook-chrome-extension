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
- Outlook Web の利用権限
  ([outlook.cloud.microsoft](https://outlook.cloud.microsoft/), [outlook.office.com](https://outlook.office.com/), [outlook.office365.com](https://outlook.office365.com/), [outlook.live.com](https://outlook.live.com/))

### 5分セットアップ（非エンジニア向け）

1. [Releases ページを開く](https://github.com/chaspy/outlook-chrome-extension/releases)
2. 最新リリースの `Assets` から `outlook-chrome-extension-vX.Y.Z.zip` をダウンロードする  
   (`Source code (zip)` ではなく、`outlook-chrome-extension-...zip` を選んでください)
3. ダウンロードした ZIP を展開する
4. Chrome で [chrome://extensions/](chrome://extensions/) を開く
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

推奨: Google スプレッドシート / Excel で「氏名・メール・ID」の3列を選択してコピーし、
そのまま貼り付けてください（タブ区切り TSV として読み取られます）。

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
- 拡張更新後に挙動が変: [chrome://extensions/](chrome://extensions/) で再読み込み

### 新しいバージョンへの更新手順

連絡先データ（`chrome.storage.local`）を消さないため、次の手順を推奨します。

1. [Releases ページ](https://github.com/chaspy/outlook-chrome-extension/releases) から最新版 ZIP をダウンロードして展開
2. 今使っている拡張フォルダに、新しいファイルを上書きコピー
3. [chrome://extensions/](chrome://extensions/) でこの拡張の `再読み込み` をクリック
4. Outlook Web を再読み込みして動作確認

上書きが難しい場合は、拡張を入れ直しても構いません。  
ただし再インストール時は連絡先データが消える可能性があります。

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
- Access to Outlook Web
  ([outlook.cloud.microsoft](https://outlook.cloud.microsoft/), [outlook.office.com](https://outlook.office.com/), [outlook.office365.com](https://outlook.office365.com/), [outlook.live.com](https://outlook.live.com/))

### 5-minute setup (for non-engineers)

1. [Open Releases](https://github.com/chaspy/outlook-chrome-extension/releases)
2. In the latest release `Assets`, download `outlook-chrome-extension-vX.Y.Z.zip`  
   (Do not choose `Source code (zip)`; choose `outlook-chrome-extension-...zip`.)
3. Extract the downloaded ZIP
4. Open [chrome://extensions/](chrome://extensions/) in Chrome
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

Recommended: In Google Sheets / Excel, select the 3 columns (`name`, `email`, `id`),
copy them, and paste directly (it will be parsed as tab-separated TSV).

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
- Strange behavior after updates: reload extension in [chrome://extensions/](chrome://extensions/)

### Updating To A New Version

To keep contact data (`chrome.storage.local`), use this flow:

1. Download and extract the latest ZIP from [Releases](https://github.com/chaspy/outlook-chrome-extension/releases)
2. Overwrite files in the folder currently loaded by Chrome
3. Click `Reload` for this extension in [chrome://extensions/](chrome://extensions/)
4. Reload Outlook Web and verify behavior

If overwriting files is difficult, reinstalling is also possible.  
Note: reinstalling may clear saved contact data.

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

### Version management (single source of truth)

`VERSION` is the canonical version value used for release tags.

```bash
# 1) update VERSION
echo "0.2.3" > VERSION

# 2) sync manifest/package/package-lock
npm run sync:version

# 3) validate for PR merge guard
npm run check:release-version
```
