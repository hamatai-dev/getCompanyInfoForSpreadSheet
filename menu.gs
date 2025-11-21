/**
 * メニュー管理用ファイル
 * すべてのカスタムメニューをここで一元管理
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("企業情報取得")
    .addSubMenu(
      ui
        .createMenu("住所から検索")
        .addItem("会社名・電話番号を取得", "searchCompanyInfoByAddress")
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("会社名から検索")
        .addItem("住所・電話番号を取得", "searchCompanyInfoByCompanyName")
    )
    .addSeparator()
    .addSubMenu(
      ui
        .createMenu("【使用不可】電話番号から検索")
        .addItem("会社名・住所を取得", "searchCompanyInfoByPhoneNumber")
    )
    .addToUi();
}
