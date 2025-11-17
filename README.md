# getCompanyInfoForSpreadSheet

## 概要 / Overview

### 日本語

この Google Apps Script は、Google スプレッドシートから住所情報を読み取り、Google Places API を使用して企業名と電話番号を自動的に取得するツールです。

### English

This Google Apps Script tool reads address information from Google Spreadsheets and automatically retrieves company names and phone numbers using the Google Places API.

---

## 主な機能 / Main Features

### 日本語

- **一括処理**: スプレッドシートの複数行を一度に処理
- **住所検索**: Google Places API (Text Search) を使用して住所から企業情報を検索
- **スコアリングシステム**: 検索結果を評価し、最も適切な企業情報を選択
  - 電話番号の有無
  - 階数情報の一致
  - 施設タイプの評価
  - 住所の一致度
- **カスタムメニュー**: スプレッドシートに専用メニューを追加
- **カスタム関数**: セル内で直接使用できる `COMPANY_INFO()` 関数

### English

- **Batch Processing**: Process multiple rows in a spreadsheet at once
- **Address Search**: Search for company information from addresses using Google Places API (Text Search)
- **Scoring System**: Evaluate search results and select the most appropriate company information
  - Phone number availability
  - Floor number matching
  - Facility type evaluation
  - Address matching accuracy
- **Custom Menu**: Add a dedicated menu to the spreadsheet
- **Custom Function**: `COMPANY_INFO()` function that can be used directly in cells

---

## 使用方法 / Usage

### 日本語

#### 初期設定

1. Google Cloud Console で新しいプロジェクトを作成
2. Places API (New) を有効化
3. API キーを取得
4. スクリプトの `API_KEY` 変数に API キーを設定

#### スプレッドシートの準備

- **A 列**: 住所を入力
- **B 列**: 企業名（自動入力される）
- **C 列**: 電話番号（自動入力される）
- **1 行目**: ヘッダー行

#### 実行方法

1. スプレッドシートを開く
2. メニューバーに表示される「企業情報取得」をクリック
3. 「住所から企業情報を取得」を選択
4. 処理が完了するまで待つ

#### カスタム関数の使用

セルに以下のように入力：

```
=COMPANY_INFO(A2)
```

### English

#### Initial Setup

1. Create a new project in Google Cloud Console
2. Enable Places API (New)
3. Obtain an API key
4. Set the API key in the `API_KEY` variable in the script

#### Spreadsheet Preparation

- **Column A**: Enter addresses
- **Column B**: Company names (auto-populated)
- **Column C**: Phone numbers (auto-populated)
- **Row 1**: Header row

#### How to Execute

1. Open the spreadsheet
2. Click "企業情報取得" (Get Company Info) in the menu bar
3. Select "住所から企業情報を取得" (Get company info from address)
4. Wait for processing to complete

#### Using Custom Function

Enter in a cell like this:

```
=COMPANY_INFO(A2)
```

---

## 関数の説明 / Function Descriptions

### 日本語

#### `main()`

- スプレッドシートの全行を処理するメイン関数
- A 列の住所を読み取り、B 列と C 列に結果を出力
- API 制限を考慮して 1 秒の待機時間を設定

#### `getCompanyInfoByAddress(address)`

- 住所から企業情報を取得する中核関数
- Google Places API の Text Search を使用
- スコアリングアルゴリズムで最適な結果を選択
- 戻り値: `[企業名, 電話番号]`

#### `calculateAddressMatch(searchAddress, resultAddress)`

- 検索住所と API 結果の一致度を計算
- 階数情報の一致を評価
- 0-30 のスコアを返す

#### `testAddress()`

- デバッグ用のテスト関数
- 特定の住所で動作確認が可能

#### `COMPANY_INFO(address)`

- スプレッドシートのセルで使用できるカスタム関数
- 企業名と電話番号を「/」で区切って返す

#### `onOpen()`

- スプレッドシート起動時に自動実行
- カスタムメニューを追加

### English

#### `main()`

- Main function to process all rows in the spreadsheet
- Reads addresses from column A and outputs results to columns B and C
- Sets a 1-second wait time considering API limits

#### `getCompanyInfoByAddress(address)`

- Core function to retrieve company information from an address
- Uses Google Places API Text Search
- Selects optimal results with a scoring algorithm
- Returns: `[company name, phone number]`

#### `calculateAddressMatch(searchAddress, resultAddress)`

- Calculates matching accuracy between search address and API results
- Evaluates floor number matching
- Returns a score of 0-30

#### `testAddress()`

- Test function for debugging
- Allows testing with specific addresses

#### `COMPANY_INFO(address)`

- Custom function that can be used in spreadsheet cells
- Returns company name and phone number separated by "/"

#### `onOpen()`

- Automatically executed when spreadsheet is opened
- Adds custom menu

---

## スコアリングロジック / Scoring Logic

### 日本語

スクリプトは以下の基準で検索結果をスコアリングします：

1. **電話番号あり**: +100 点
2. **階数情報の一致**: +50 点
3. **施設名のみで電話なし**: -50 点
4. **観光地・施設のみ**: -30 点
5. **住所の一致度**: 0-30 点

### English

The script scores search results based on the following criteria:

1. **Phone number available**: +100 points
2. **Floor number match**: +50 points
3. **Facility name only without phone**: -50 points
4. **Tourist spot/facility only**: -30 points
5. **Address matching accuracy**: 0-30 points

---

## 注意事項 / Notes

### 日本語

- Google Places API の使用には料金が発生する場合があります
- API キーは安全に管理してください
- API 制限に注意してください（デフォルトで 1 秒の待機時間を設定）
- 検索結果が常に正確とは限りません

### English

- Google Places API usage may incur charges
- Keep your API key secure
- Be mindful of API limits (default 1-second wait time is set)
- Search results may not always be accurate

---

## 必要な権限 / Required Permissions

### 日本語

- Google スプレッドシートへのアクセス
- 外部サービスへの接続（Google Places API）

### English

- Access to Google Spreadsheets
- Connection to external services (Google Places API)
