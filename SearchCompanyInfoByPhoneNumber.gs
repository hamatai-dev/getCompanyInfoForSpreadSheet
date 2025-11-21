/**
 * 複数行を一括処理する関数（電話番号検索用）
 */
function searchCompanyInfoByPhoneNumber() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  // A列に電話番号、B列に会社名、C列に住所を出力する想定
  const phoneColumn = 1; // A列
  const companyNameColumn = 2; // B列
  const addressColumn = 3; // C列
  const startRow = 2; // 2行目から開始（1行目はヘッダー）

  for (let i = startRow; i <= lastRow; i++) {
    const phoneNumber = sheet.getRange(i, phoneColumn).getValue();

    if (phoneNumber && phoneNumber !== "") {
      Logger.log(`\n処理中: 行${i} - ${phoneNumber}`);
      const [companyName, address] = getCompanyInfoByPhone(phoneNumber);

      sheet.getRange(i, companyNameColumn).setValue(companyName);
      sheet.getRange(i, addressColumn).setValue(address);

      // API制限を考慮して少し待機
      Utilities.sleep(1000);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("処理完了", "完了", 3);
}

/**
 * 電話番号から企業情報を取得するメイン関数（新Places API対応版）
 * @param {string} phoneNumber - 検索する電話番号
 * @return {Array} [会社名, 住所]
 */
function getCompanyInfoByPhone(phoneNumber) {
  if (!phoneNumber || phoneNumber === "") {
    return ["", ""];
  }

  try {
    // 電話番号を国際形式に変換（日本の場合）
    const formattedPhone = formatPhoneForAPI(phoneNumber);

    // 新しいPlaces API (Text Search)を使用
    const searchUrl = "https://places.googleapis.com/v1/places:searchText";

    const payload = {
      textQuery: formattedPhone,
      languageCode: "ja",
      regionCode: "JP",
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: {
        "X-Goog-Api-Key": API_KEY,
        "X-Goog-FieldMask":
          "places.displayName,places.formattedAddress,places.internationalPhoneNumber,places.nationalPhoneNumber,places.id,places.types,places.businessStatus",
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    Logger.log("元の電話番号: " + phoneNumber);
    Logger.log("API用フォーマット: " + formattedPhone);
    Logger.log("検索ペイロード: " + JSON.stringify(payload));

    const response = UrlFetchApp.fetch(searchUrl, options);
    const responseCode = response.getResponseCode();
    const data = JSON.parse(response.getContentText());

    Logger.log("HTTPステータス: " + responseCode);

    if (responseCode === 403 || responseCode === 400) {
      Logger.log("エラー詳細: " + JSON.stringify(data));
      return ["APIエラー: " + (data.error ? data.error.message : "不明"), ""];
    }

    if (data.places && data.places.length > 0) {
      Logger.log("検索結果数: " + data.places.length);

      // 最適な結果を選択するためのスコアリング
      let bestPlace = null;
      let bestScore = -1;

      for (let i = 0; i < data.places.length; i++) {
        const place = data.places[i];
        let score = 0;

        const placeName = place.displayName ? place.displayName.text : "";
        const formattedAddress = place.formattedAddress || "";
        const placePhone =
          place.nationalPhoneNumber || place.internationalPhoneNumber || "";
        const types = place.types || [];
        const businessStatus = place.businessStatus || "";

        Logger.log(`\n結果${i + 1}: ${placeName}`);
        Logger.log(`  住所: ${formattedAddress}`);
        Logger.log(`  電話: ${placePhone || "電話なし"}`);
        Logger.log(`  タイプ: ${types.join(", ")}`);
        Logger.log(`  営業状態: ${businessStatus}`);

        // スコアリングロジック

        // 1. 電話番号の一致度（最重要）
        const phoneMatch = calculatePhoneMatch(phoneNumber, placePhone);
        score += phoneMatch;
        Logger.log(`  [+${phoneMatch}] 電話番号一致度`);

        // 電話番号が一致しない場合はスキップ
        if (phoneMatch < 100) {
          Logger.log(`  スキップ: 電話番号が十分に一致しません`);
          continue;
        }

        // 2. 会社名があれば高スコア
        if (placeName) {
          score += 50;
          Logger.log(`  [+50] 会社名あり`);
        }

        // 3. 営業状態によるスコア調整
        if (businessStatus === "OPERATIONAL") {
          score += 30;
          Logger.log(`  [+30] 営業中`);
        } else if (businessStatus === "CLOSED_TEMPORARILY") {
          score += 10;
          Logger.log(`  [+10] 一時休業中`);
        } else if (businessStatus === "CLOSED_PERMANENTLY") {
          score -= 50;
          Logger.log(`  [-50] 永久閉店`);
        }

        // 4. 住所が具体的であれば高スコア
        if (formattedAddress && formattedAddress.length > 0) {
          score += 20;
          Logger.log(`  [+20] 住所あり`);
        }

        // 5. ビジネス関連のタイプであれば高スコア
        const businessTypes = [
          "establishment",
          "point_of_interest",
          "store",
          "restaurant",
          "food",
        ];
        const hasBusinessType = types.some((type) =>
          businessTypes.includes(type)
        );
        if (hasBusinessType) {
          score += 10;
          Logger.log(`  [+10] ビジネスタイプ該当`);
        }

        // 6. 地名や観光地のみの場合は低スコア
        const locationOnlyTypes = [
          "locality",
          "political",
          "tourist_attraction",
        ];
        const isLocationOnly =
          types.length > 0 &&
          types.every((type) => locationOnlyTypes.includes(type));
        if (isLocationOnly) {
          score -= 50;
          Logger.log(`  [-50] 地名・観光地のみ`);
        }

        Logger.log(`  合計スコア: ${score}`);

        if (score > bestScore) {
          bestScore = score;
          bestPlace = place;
        }
      }

      if (bestPlace && bestScore > 0) {
        const companyName = bestPlace.displayName
          ? bestPlace.displayName.text
          : "検索結果なし";
        const rawAddress = bestPlace.formattedAddress || "検索結果なし";

        // 住所を整形
        const formattedAddress = formatAddress(rawAddress);

        Logger.log("\n=== 選択した結果 ===");
        Logger.log("会社名: " + companyName);
        Logger.log("元の住所: " + rawAddress);
        Logger.log("整形後の住所: " + formattedAddress);
        Logger.log("スコア: " + bestScore);

        return [companyName, formattedAddress];
      } else {
        return ["検索結果なし", "検索結果なし"];
      }
    } else {
      Logger.log("検索結果が0件でした");
      return ["検索結果なし", "検索結果なし"];
    }
  } catch (error) {
    Logger.log("エラー: " + error.toString());
    return ["エラー: " + error.toString(), ""];
  }
}

/**
 * 電話番号をAPI検索用にフォーマットする関数
 * Google Places APIの推奨形式：国コード + スペース + 電話番号
 * 例: 03-1234-5678 → +81 3-1234-5678
 * 例: 092-555-1234 → +81 92-555-1234
 * @param {string} phoneNumber - 元の電話番号
 * @return {string} フォーマット済み電話番号
 */
function formatPhoneForAPI(phoneNumber) {
  if (!phoneNumber) return "";

  let formatted = String(phoneNumber).trim();

  // 全角数字を半角に変換
  formatted = formatted.replace(/[０-９]/g, (s) => {
    return String.fromCharCode(s.charCodeAt(0) - 0xfee0);
  });

  // 既に+81で始まっている場合はそのまま返す
  if (formatted.startsWith("+81")) {
    // +81の後にスペースがない場合は追加
    if (!formatted.startsWith("+81 ")) {
      formatted = formatted.replace(/^\+81/, "+81 ");
    }
    return formatted;
  }

  // 先頭の0を削除（日本国内の電話番号は0で始まる）
  if (formatted.startsWith("0")) {
    formatted = formatted.substring(1);
  }

  // 国コード(+81)とスペースを追加
  formatted = "+81 " + formatted;

  return formatted;
}

/**
 * 電話番号を正規化する関数（比較用）
 * ハイフン、括弧、スペース、プラス記号を削除し、数字のみにする
 * @param {string} phoneNumber - 元の電話番号
 * @return {string} 正規化された電話番号
 */
function normalizePhoneNumber(phoneNumber) {
  if (!phoneNumber) return "";

  let normalized = String(phoneNumber);

  // 全角数字を半角に変換
  normalized = normalized.replace(/[０-９]/g, (s) => {
    return String.fromCharCode(s.charCodeAt(0) - 0xfee0);
  });

  // ハイフン、括弧、スペース、プラス記号を削除
  normalized = normalized.replace(/[-\(\)\s+]/g, "");

  // 国コード81を0に変換（日本国内形式に統一）
  if (normalized.startsWith("81")) {
    normalized = "0" + normalized.substring(2);
  }

  return normalized;
}

/**
 * 電話番号の一致度を計算
 * @param {string} searchPhone - 検索した電話番号
 * @param {string} resultPhone - APIから返された電話番号
 * @return {number} 一致度スコア（0-150）
 */
function calculatePhoneMatch(searchPhone, resultPhone) {
  if (!searchPhone || !resultPhone) return 0;

  // 両方を正規化して比較
  const normalizedSearch = normalizePhoneNumber(searchPhone);
  const normalizedResult = normalizePhoneNumber(resultPhone);

  Logger.log(
    `  電話番号比較（正規化後）: "${normalizedSearch}" vs "${normalizedResult}"`
  );

  let score = 0;

  // 完全一致（最高スコア）
  if (normalizedSearch === normalizedResult) {
    score = 150;
  }
  // 末尾が一致（市外局番なしで検索された場合などに対応）
  else if (
    normalizedResult.endsWith(normalizedSearch) ||
    normalizedSearch.endsWith(normalizedResult)
  ) {
    const endMatchLength = Math.min(
      normalizedSearch.length,
      normalizedResult.length
    );
    if (endMatchLength >= 8) {
      // 少なくとも8桁一致
      score = 120;
    } else if (endMatchLength >= 6) {
      // 6-7桁一致
      score = 80;
    }
  }
  // 部分一致
  else if (
    normalizedResult.includes(normalizedSearch) ||
    normalizedSearch.includes(normalizedResult)
  ) {
    score = 60;
  }

  return score;
}

/**
 * 住所を整形する関数
 * 「日本、〒xxx-xxxx」の部分を削除し、都道府県から始まる住所のみを返す
 * @param {string} address - APIから取得した住所
 * @return {string} 整形後の住所
 */
function formatAddress(address) {
  if (!address || address === "") {
    return "";
  }

  // 「日本、」を削除
  let formatted = address.replace(/^日本、?\s*/g, "");

  // 郵便番号（〒xxx-xxxx）を削除
  formatted = formatted.replace(/〒\d{3}-?\d{4}\s*/g, "");

  // 先頭と末尾の空白を削除
  formatted = formatted.trim();

  return formatted;
}

/**
 * デバック用のテスト関数 - 特定の電話番号で動作確認
 */
function testPhoneNumber() {
  // テスト用の電話番号（実際に存在する電話番号に変更してください）
  const testPhoneNumber = "092-406-0095"; // 例: 東京都庁の代表番号

  Logger.log("=== テスト開始（電話番号検索） ===");
  const result = getCompanyInfoByPhone(testPhoneNumber);

  Logger.log("\n=== 最終結果 ===");
  Logger.log("会社名: " + result[0]);
  Logger.log("住所: " + result[1]);
}

/**
 * スプレッドシートで使用するカスタム関数
 * セルに「=COMPANY_INFO_BY_PHONE(A2)」のように入力して使用
 */
function COMPANY_INFO_BY_PHONE(phoneNumber) {
  const info = getCompanyInfoByPhone(phoneNumber);
  return info[0] + " / " + info[1];
}
