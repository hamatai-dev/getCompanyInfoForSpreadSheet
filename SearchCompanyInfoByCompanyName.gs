/**
 * 複数行を一括処理する関数
 */
function searchCompanyInfoByCompanyName() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  // A列に会社名、B列に住所、C列に電話番号を出力する想定
  const companyNameColumn = 1; // A列
  const addressColumn = 2; // B列
  const phoneColumn = 3; // C列
  const startRow = 2; // 2行目から開始（1行目はヘッダー）

  for (let i = startRow; i <= lastRow; i++) {
    const companyName = sheet.getRange(i, companyNameColumn).getValue();

    if (companyName && companyName !== "") {
      Logger.log(`\n処理中: 行${i} - ${companyName}`);
      const [address, phoneNumber] = getCompanyInfoByName(companyName);

      sheet.getRange(i, addressColumn).setValue(address);
      sheet.getRange(i, phoneColumn).setValue(phoneNumber);

      // API制限を考慮して少し待機
      Utilities.sleep(1000);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("処理完了", "完了", 3);
}

/**
 * 会社名から企業情報を取得するメイン関数（新Places API対応版）
 * @param {string} companyName - 検索する会社名
 * @return {Array} [住所, 電話番号]
 */
function getCompanyInfoByName(companyName) {
  if (!companyName || companyName === "") {
    return ["", ""];
  }

  try {
    // 新しいPlaces API (Text Search)を使用
    const searchUrl = "https://places.googleapis.com/v1/places:searchText";

    const payload = {
      textQuery: companyName,
      languageCode: "ja",
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

    Logger.log("検索会社名: " + companyName);

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
        const phoneNumber =
          place.nationalPhoneNumber || place.internationalPhoneNumber || "";
        const types = place.types || [];
        const businessStatus = place.businessStatus || "";

        Logger.log(`\n結果${i + 1}: ${placeName}`);
        Logger.log(`  住所: ${formattedAddress}`);
        Logger.log(`  電話: ${phoneNumber || "電話なし"}`);
        Logger.log(`  タイプ: ${types.join(", ")}`);
        Logger.log(`  営業状態: ${businessStatus}`);

        // スコアリングロジック

        // 1. 会社名の一致度（最重要）
        const nameMatch = calculateNameMatch(companyName, placeName);
        score += nameMatch;
        Logger.log(`  [+${nameMatch}] 会社名一致度`);

        // 2. 電話番号があれば高スコア
        if (phoneNumber) {
          score += 50;
          Logger.log(`  [+50] 電話番号あり`);
        }

        // 3. 営業中の企業であれば高スコア
        if (businessStatus === "OPERATIONAL") {
          score += 30;
          Logger.log(`  [+30] 営業中`);
        }

        // 4. 住所が具体的であれば高スコア
        if (formattedAddress && formattedAddress.includes("日本")) {
          score += 20;
          Logger.log(`  [+20] 日本の住所あり`);
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
        const isLocationOnly = types.every((type) =>
          locationOnlyTypes.includes(type)
        );
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
        const rawAddress = bestPlace.formattedAddress || "検索結果なし";
        const phoneNumber =
          bestPlace.nationalPhoneNumber ||
          bestPlace.internationalPhoneNumber ||
          "検索結果なし";

        // 住所を整形
        const formattedAddress = formatAddress(rawAddress);

        Logger.log("\n=== 選択した結果 ===");
        Logger.log("元の住所: " + rawAddress);
        Logger.log("整形後の住所: " + formattedAddress);
        Logger.log("電話番号: " + phoneNumber);
        Logger.log("スコア: " + bestScore);

        return [formattedAddress, phoneNumber];
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
 * 会社名の一致度を計算
 * @param {string} searchName - 検索した会社名
 * @param {string} resultName - APIから返された会社名
 * @return {number} 一致度スコア（0-100）
 */
function calculateNameMatch(searchName, resultName) {
  if (!searchName || !resultName) return 0;

  // 正規化（空白削除、小文字化、株式会社などの除去）
  const normalize = (str) => {
    return str
      .replace(/\s+/g, "") // 空白を削除
      .replace(/株式会社|有限会社|合同会社|（株）|\(株\)|（有）|\(有\)/g, "") // 法人格を削除
      .toLowerCase();
  };

  const normalizedSearch = normalize(searchName);
  const normalizedResult = normalize(resultName);

  let score = 0;

  // 完全一致
  if (normalizedSearch === normalizedResult) {
    score = 100;
  }
  // 部分一致（検索語が結果に含まれる）
  else if (normalizedResult.includes(normalizedSearch)) {
    score = 80;
  }
  // 部分一致（結果が検索語に含まれる）
  else if (normalizedSearch.includes(normalizedResult)) {
    score = 70;
  }
  // 前方一致
  else if (
    normalizedResult.startsWith(normalizedSearch) ||
    normalizedSearch.startsWith(normalizedResult)
  ) {
    score = 60;
  }

  return score;
}

/**
 * デバック用のテスト関数 - 特定の会社名で動作確認
 */
function testCompanyName() {
  const testCompanyName = "アルサーガパートナーズ 福岡";

  Logger.log("=== テスト開始 ===");
  const result = getCompanyInfoByName(testCompanyName);

  Logger.log("\n=== 最終結果 ===");
  Logger.log("住所: " + result[0]);
  Logger.log("電話番号: " + result[1]);
}
