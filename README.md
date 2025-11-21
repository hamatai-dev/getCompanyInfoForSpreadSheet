# getCompanyInfoForSpreadSheet

## 概要 / Overview

### 日本語

この Google Apps Script は、Google スプレッドシートから住所や会社名の情報を読み取り、Google Places API を使用して企業情報を自動的に取得するツールです。住所から会社名・電話番号を検索する機能と、会社名から住所・電話番号を検索する機能の両方を提供します。

### English

This Google Apps Script tool reads address or company name information from Google Spreadsheets and automatically retrieves company information using the Google Places API. It provides both features to search for company names and phone numbers from addresses, and to search for addresses and phone numbers from company names.

---

---

## 主な機能 / Main Features

### 日本語

- **一括処理**: スプレッドシートの複数行を一度に処理
- **住所検索**: Google Places API (Text Search) を使用して住所から会社名と電話番号を検索
- **会社名検索**: Google Places API (Text Search) を使用して会社名から住所と電話番号を検索
- **スコアリングシステム**: 検索結果を評価し、最も適切な企業情報を選択
  - **住所検索時**:
    - 電話番号の有無
    - 階数情報の一致
    - 施設タイプの評価
    - 住所の一致度
  - **会社名検索時**:
    - 会社名の一致度（最重要）
    - 電話番号の有無
    - 営業状態（OPERATIONAL）
    - 住所の具体性
    - ビジネスタイプの該当
- **カスタムメニュー**: スプレッドシートに専用メニューを追加（住所検索と会社名検索を選択可能）
- **モジュール構造**: コードを機能別に 4 つのファイルに分割（config.gs, menu.gs, SearchCompanyInfoByAddress.gs, SearchCompanyInfoByCompanyName.gs）

### English

- **Batch Processing**: Process multiple rows in a spreadsheet at once
- **Address Search**: Search for company names and phone numbers from addresses using Google Places API (Text Search)
- **Company Name Search**: Search for addresses and phone numbers from company names using Google Places API (Text Search)
- **Scoring System**: Evaluate search results and select the most appropriate company information
  - **For Address Search**:
    - Phone number availability
    - Floor number matching
    - Facility type evaluation
    - Address matching accuracy
  - **For Company Name Search**:
    - Company name matching accuracy (most important)
    - Phone number availability
    - Business status (OPERATIONAL)
    - Address specificity
    - Business type relevance
- **Custom Menu**: Add a dedicated menu to the spreadsheet (selectable between address search and company name search)
- **Modular Structure**: Code split into 4 files by function (config.gs, menu.gs, SearchCompanyInfoByAddress.gs, SearchCompanyInfoByCompanyName.gs)

---

## 使用方法 / Usage

### 日本語

#### 初期設定

1. Google Cloud Console で新しいプロジェクトを作成
2. Places API (New) を有効化
3. API キーを取得
4. スクリプトの `API_KEY` 変数に API キーを設定

#### スプレッドシートの準備

**住所から会社名・電話番号を検索する場合:**

- **A 列**: 住所を入力
- **B 列**: 会社名（自動入力される）
- **C 列**: 電話番号（自動入力される）
- **1 行目**: ヘッダー行

**会社名から住所・電話番号を検索する場合:**

- **A 列**: 会社名を入力
- **B 列**: 住所（自動入力される）
- **C 列**: 電話番号（自動入力される）
- **1 行目**: ヘッダー行

#### 実行方法

1. スプレッドシートを開く
2. メニューバーに表示される「企業情報取得」をクリック
3. 以下のいずれかを選択：
   - 「住所から検索」→「会社名・電話番号を取得」
   - 「会社名から検索」→「住所・電話番号を取得」
4. 処理が完了するまで待つ

### English

#### Initial Setup

1. Create a new project in Google Cloud Console
2. Enable Places API (New)
3. Obtain an API key
4. Set the API key in the `API_KEY` variable in the `config.gs` file

#### Spreadsheet Preparation

**To search for company name and phone number from address:**

- **Column A**: Enter addresses
- **Column B**: Company names (auto-populated)
- **Column C**: Phone numbers (auto-populated)
- **Row 1**: Header row

**To search for address and phone number from company name:**

- **Column A**: Enter company names
- **Column B**: Addresses (auto-populated)
- **Column C**: Phone numbers (auto-populated)
- **Row 1**: Header row

#### How to Execute

1. Open the spreadsheet
2. Click "企業情報取得" (Get Company Info) in the menu bar
3. Select one of the following:
   - "住所から検索" (Search from address) → "会社名・電話番号を取得" (Get company name and phone number)
   - "会社名から検索" (Search from company name) → "住所・電話番号を取得" (Get address and phone number)
4. Wait for processing to complete

---

## ファイル構成 / File Structure

### 日本語

- **config.gs**: API キーの設定
- **menu.gs**: カスタムメニューの管理
- **SearchCompanyInfoByAddress.gs**: 住所から会社名・電話番号を検索する機能
- **SearchCompanyInfoByCompanyName.gs**: 会社名から住所・電話番号を検索する機能

### English

- **config.gs**: API key configuration
- **menu.gs**: Custom menu management
- **SearchCompanyInfoByAddress.gs**: Feature to search for company name and phone number from address
- **SearchCompanyInfoByCompanyName.gs**: Feature to search for address and phone number from company name

---

## 関数の説明 / Function Descriptions

### 日本語

#### SearchCompanyInfoByAddress.gs

##### `searchCompanyInfoByAddress()`

- スプレッドシートの全行を処理するメイン関数
- A 列の住所を読み取り、B 列（会社名）と C 列（電話番号）に結果を出力
- API 制限を考慮して 1 秒の待機時間を設定

##### `getCompanyInfoByAddress(address)`

- 住所から企業情報を取得する中核関数
- Google Places API の Text Search を使用
- スコアリングアルゴリズムで最適な結果を選択
- 戻り値: `[企業名, 電話番号]`

##### `calculateAddressMatch(searchAddress, resultAddress)`

- 検索住所と API 結果の一致度を計算
- 階数情報の一致を評価
- 0-30 のスコアを返す

#### SearchCompanyInfoByCompanyName.gs

##### `searchCompanyInfoByCompanyName()`

- スプレッドシートの全行を処理するメイン関数
- A 列の会社名を読み取り、B 列（住所）と C 列（電話番号）に結果を出力
- API 制限を考慮して 1 秒の待機時間を設定

##### `getCompanyInfoByName(companyName)`

- 会社名から企業情報を取得する中核関数
- Google Places API の Text Search を使用
- スコアリングアルゴリズムで最適な結果を選択
- 戻り値: `[住所, 電話番号]`

##### `calculateNameMatch(searchName, resultName)`

- 検索会社名と API 結果の一致度を計算
- 会社名の正規化（法人格の除去など）を行い、一致度を評価
- 0-100 のスコアを返す（完全一致=100、部分一致=80、前方一致=60 など）

##### `formatAddress(address)`

- API から取得した住所を整形
- 「日本、」や郵便番号（〒 xxx-xxxx）を削除し、都道府県から始まる住所のみを返す

##### `testCompanyName()`

- デバッグ用のテスト関数
- 特定の会社名で動作確認が可能

#### menu.gs

##### `onOpen()`

- スプレッドシート起動時に自動実行
- カスタムメニューを追加（住所検索と会社名検索のサブメニュー付き）

### English

#### SearchCompanyInfoByAddress.gs

##### `searchCompanyInfoByAddress()`

- Main function to process all rows in the spreadsheet
- Reads addresses from column A and outputs results to columns B (company name) and C (phone number)
- Sets a 1-second wait time considering API limits

##### `getCompanyInfoByAddress(address)`

- Core function to retrieve company information from an address
- Uses Google Places API Text Search
- Selects optimal results with a scoring algorithm
- Returns: `[company name, phone number]`

##### `calculateAddressMatch(searchAddress, resultAddress)`

- Calculates matching accuracy between search address and API results
- Evaluates floor number matching
- Returns a score of 0-30

#### SearchCompanyInfoByCompanyName.gs

##### `searchCompanyInfoByCompanyName()`

- Main function to process all rows in the spreadsheet
- Reads company names from column A and outputs results to columns B (address) and C (phone number)
- Sets a 1-second wait time considering API limits

##### `getCompanyInfoByName(companyName)`

- Core function to retrieve company information from a company name
- Uses Google Places API Text Search
- Selects optimal results with a scoring algorithm
- Returns: `[address, phone number]`

##### `calculateNameMatch(searchName, resultName)`

- Calculates matching accuracy between search company name and API results
- Normalizes company names (removes corporate types, etc.) and evaluates matching accuracy
- Returns a score of 0-100 (exact match=100, partial match=80, prefix match=60, etc.)

##### `formatAddress(address)`

- Formats addresses retrieved from API
- Removes "Japan," and postal codes (〒 xxx-xxxx), returning only addresses starting from prefecture

##### `testCompanyName()`

- Test function for debugging
- Allows testing with specific company names

#### menu.gs

##### `onOpen()`

- Automatically executed when spreadsheet is opened
- Adds custom menu with submenus for address search and company name search

---

## スコアリングロジック / Scoring Logic

### 日本語（住所検索）

住所から会社名を検索する際、以下の基準で検索結果をスコアリングします：

1. **電話番号あり**: +100 点
2. **階数情報の一致**: +50 点
3. **施設名のみで電話なし**: -50 点
4. **観光地・施設のみ**: -30 点
5. **住所の一致度**: 0-30 点

### 日本語（会社名検索）

会社名から住所を検索する際、以下の基準で検索結果をスコアリングします：

1. **会社名の一致度**: 0-100 点（完全一致=100、部分一致=80、前方一致=60 など）
2. **電話番号あり**: +50 点
3. **営業中（OPERATIONAL）**: +30 点
4. **日本の住所あり**: +20 点
5. **ビジネスタイプ該当**: +10 点
6. **地名・観光地のみ**: -50 点

### English (Address Search)

When searching for company names from addresses, the script scores search results based on the following criteria:

1. **Phone number available**: +100 points
2. **Floor number match**: +50 points
3. **Facility name only without phone**: -50 points
4. **Tourist spot/facility only**: -30 points
5. **Address matching accuracy**: 0-30 points

### English (Company Name Search)

When searching for addresses from company names, the script scores search results based on the following criteria:

1. **Company name matching accuracy**: 0-100 points (exact match=100, partial match=80, prefix match=60, etc.)
2. **Phone number available**: +50 points
3. **Business status OPERATIONAL**: +30 points
4. **Japanese address available**: +20 points
5. **Business type relevance**: +10 points
6. **Location/tourist attraction only**: -50 points

---

## 注意事項 / Important Notes

### 日本語の注意事項

- Google Places API の使用には料金が発生する場合があります
- API キーは `config.gs` ファイルに設定し、安全に管理してください
- API 制限に注意してください（デフォルトで 1 秒の待機時間を設定）
- 検索結果が常に正確とは限りません
- 会社名検索では、検索語に地域名（例：「福岡」）を含めるとより精度の高い結果が得られます
- 会社名検索で取得した住所は、「日本、」や郵便番号が自動的に削除され、都道府県から始まる形式で返されます

### Notes in English

- Google Places API usage may incur charges
- Set your API key in the `config.gs` file and keep it secure
- Be mindful of API limits (default 1-second wait time is set)
- Search results may not always be accurate
- For company name searches, including location names (e.g., "Fukuoka") in the search query can improve result accuracy
- Addresses retrieved from company name searches automatically have "Japan," and postal codes removed, returning addresses starting from prefecture

---

## 必要な権限 / Required Permissions

### 日本語の必要な権限

- Google スプレッドシートへのアクセス
- 外部サービスへの接続（Google Places API）

### Permissions Required

- Access to Google Spreadsheets
- Connection to external services (Google Places API)
