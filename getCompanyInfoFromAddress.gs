// Google Places APIキーを設定（Google Cloud Consoleで取得）
const API_KEY = "Google Places APIキー";

/**
 * 複数行を一括処理する関数
 */
function main() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();

  // A列に住所、B列に企業名、C列に電話番号を出力する想定
  const addressColumn = 1; // A列
  const nameColumn = 2; // B列
  const phoneColumn = 3; // C列
  const startRow = 2; // 2行目から開始（1行目はヘッダー）

  for (let i = startRow; i <= lastRow; i++) {
    const address = sheet.getRange(i, addressColumn).getValue();

    if (address && address !== "") {
      Logger.log(`\n処理中: 行${i} - ${address}`);
      const [companyName, phoneNumber] = getCompanyInfoByAddress(address);

      sheet.getRange(i, nameColumn).setValue(companyName);
      sheet.getRange(i, phoneColumn).setValue(phoneNumber);

      // API制限を考慮して少し待機
      Utilities.sleep(1000);
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("処理完了", "完了", 3);
}

/**
 * 住所から企業情報を取得するメイン関数（新Places API対応版）
 * @param {string} address - 検索する住所
 * @return {Array} [企業名, 電話番号]
 */
function getCompanyInfoByAddress(address) {
  if (!address || address === "") {
    return ["", ""];
  }

  try {
    // 新しいPlaces API (Text Search)を使用
    const searchUrl = "https://places.googleapis.com/v1/places:searchText";

    const payload = {
      textQuery: address,
      languageCode: "ja",
    };

    const options = {
      method: "post",
      contentType: "application/json",
      headers: {
        "X-Goog-Api-Key": API_KEY,
        "X-Goog-FieldMask":
          "places.displayName,places.formattedAddress,places.internationalPhoneNumber,places.nationalPhoneNumber,places.id,places.types",
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };

    Logger.log("検索住所: " + address);

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

        const companyName = place.displayName ? place.displayName.text : "";
        const formattedAddress = place.formattedAddress || "";
        const phoneNumber =
          place.nationalPhoneNumber || place.internationalPhoneNumber || "";
        const types = place.types || [];

        Logger.log(`\n結果${i + 1}: ${companyName}`);
        Logger.log(`  住所: ${formattedAddress}`);
        Logger.log(`  電話: ${phoneNumber || "電話なし"}`);
        Logger.log(`  タイプ: ${types.join(", ")}`);

        // スコアリングロジック

        // 1. 電話番号があれば高スコア（企業である可能性が高い）
        if (phoneNumber) {
          score += 100;
          Logger.log(`  [+100] 電話番号あり`);
        }

        // 2. 階数が含まれている住所であれば高スコア
        if (formattedAddress.includes("階") || /\d+F/i.test(formattedAddress)) {
          score += 50;
          Logger.log(`  [+50] 階数情報あり`);
        }

        // 3. ビル名や施設名のみの場合は低スコア
        if (types.includes("premise") || types.includes("establishment")) {
          if (!phoneNumber) {
            score -= 50;
            Logger.log(`  [-50] 施設名のみで電話なし`);
          }
        }

        // 4. point_of_interest（観光地・施設）タイプのみの場合は低スコア
        if (types.includes("point_of_interest") && types.length === 1) {
          score -= 30;
          Logger.log(`  [-30] 観光地・施設のみ`);
        }

        // 5. 検索住所との一致度
        const addressMatch = calculateAddressMatch(address, formattedAddress);
        score += addressMatch;
        Logger.log(`  [+${addressMatch}] 住所一致度`);

        Logger.log(`  合計スコア: ${score}`);

        if (score > bestScore) {
          bestScore = score;
          bestPlace = place;
        }
      }

      if (bestPlace) {
        const companyName = bestPlace.displayName
          ? bestPlace.displayName.text
          : "";
        const phoneNumber =
          bestPlace.nationalPhoneNumber ||
          bestPlace.internationalPhoneNumber ||
          "";

        Logger.log("\n=== 選択した結果 ===");
        Logger.log("企業名: " + companyName);
        Logger.log("電話番号: " + phoneNumber);
        Logger.log("スコア: " + bestScore);

        return [companyName, phoneNumber];
      } else {
        return ["検索結果なし", ""];
      }
    } else {
      Logger.log("検索結果が0件でした");
      return ["検索結果なし", ""];
    }
  } catch (error) {
    Logger.log("エラー: " + error.toString());
    return ["エラー: " + error.toString(), ""];
  }
}

/**
 * 住所の一致度を計算
 * @param {string} searchAddress - 検索した住所
 * @param {string} resultAddress - APIから返された住所
 * @return {number} 一致度スコア（0-30）
 */
function calculateAddressMatch(searchAddress, resultAddress) {
  if (!searchAddress || !resultAddress) return 0;

  // 数字とハイフンを正規化
  const normalize = (str) => {
    return str
      .replace(/[０-９]/g, (s) => String.fromCharCode(s.charCodeAt(0) - 0xfee0)) // 全角数字を半角に
      .replace(/[−ー]/g, "-") // 各種ハイフンを統一
      .replace(/\s+/g, ""); // 空白を削除
  };

  const normalizedSearch = normalize(searchAddress);
  const normalizedResult = normalize(resultAddress);

  let score = 0;

  // 階数が含まれているか
  const floorMatch =
    normalizedSearch.match(/(\d+)階/) || normalizedSearch.match(/(\d+)F/i);
  if (floorMatch) {
    const floor = floorMatch[1];
    if (
      normalizedResult.includes(floor + "階") ||
      normalizedResult.includes(floor + "F")
    ) {
      score += 30;
    }
  }

  return score;
}

/**
 * デバック用のテスト関数 - 特定の住所で動作確認
 */
function testAddress() {
  const testAddress = "検索したい住所を指定";

  Logger.log("=== テスト開始 ===");
  const result = getCompanyInfoByAddress(testAddress);

  Logger.log("\n=== 最終結果 ===");
  Logger.log("企業名: " + result[0]);
  Logger.log("電話番号: " + result[1]);
}

/**
 * スプレッドシートで使用するカスタム関数
 * セルに「=COMPANY_INFO(A2)」のように入力して使用
 */
function COMPANY_INFO(address) {
  const info = getCompanyInfoByAddress(address);
  return info[0] + " / " + info[1];
}

/**
 * メニューにカスタム関数を追加
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("企業情報取得")
    .addItem("住所から企業情報を取得", "main")
    // .addItem('テスト実行', 'testAddress')
    .addToUi();
}
