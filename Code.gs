/**
 * 新入生名簿（Googleフォームの回答シート）の1レコードを表示・編集するWebアプリ
 *
 * 前提: 対象シートの1行目が見出し（フォームの質問文）、2行目以降が回答
 */

// ===== 基本設定 =====
const SPREADSHEET_ID = '1ao7fJBqOoEwF9ySW2cUQlhnD03GkQbhxQdrp3Un2DW0';
const SHEET_NAME = 'フォームの回答 1';
const HEADER_ROW = 1;

// ===== 列ごとの設定（見出しの文字列で指定。見出しを変えたらここも直す） =====

// 一覧表に表示する列（それ以外は編集画面でだけ表示）
const LIST_COLUMNS = [
  '氏名（大学登録表記）',
  '氏名（よみかた）',
  '学籍番号・職員番号',
  '入学年',
  '入学月',
  '研究室進学時の課程',
];

// 編集不可にする列
const READONLY_COLUMNS = ['タイムスタンプ'];

// 日付入力（カレンダー）にする列
const DATE_COLUMNS = ['生年月日'];

// 文字列として保存する列（先頭の0が消えたり、数値に変換されたりするのを防ぐ）
// 見出しにこの語が含まれる列が対象
const TEXT_COLUMN_KEYWORDS = ['電話番号', '郵便番号', '学籍番号', '会員番号', '研究者番号'];

// 入力候補を出す列（既存の回答 + ここで指定した候補から選べる。自由入力も可）
const SUGGESTIONS = {
  '入学月': ['4(Apr)', '10(Oct)'],
  '研究室進学時の課程': ['学部4年', '修士1年', '博士1年'],
  '作業着色': [],
  '債主登録': ['債主登録済'],
};

// 見出しが長すぎる列の表示名（見出しにキーの語が含まれていれば置き換える）
const SHORT_LABELS = {
  '債主登録': '債主登録（済ならチェック）',
};

/** Webアプリとしてアクセスされたときに画面を返す */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('建築材料学研究室 名簿の表示・編集')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** 対象シートを取得 */
function getSheet_() {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('シート「' + SHEET_NAME + '」が見つかりません');
  return sheet;
}

/** 見出しにキーワードが含まれるか（SUGGESTIONS / SHORT_LABELS 用） */
function findKey_(obj, header) {
  return Object.keys(obj).find(function (k) { return header.indexOf(k) !== -1; });
}

/** 列ごとの表示名・入力タイプなどをまとめる */
function buildColumns_(headers, rows) {
  return headers.map(function (h, i) {
    const labelKey = findKey_(SHORT_LABELS, h);
    const suggestKey = findKey_(SUGGESTIONS, h);
    let suggestions = null;
    if (suggestKey) {
      const set = new Set(SUGGESTIONS[suggestKey]);
      rows.forEach(function (r) { if (r.values[i] !== '') set.add(r.values[i]); });
      suggestions = Array.from(set);
    }
    return {
      header: h,
      label: labelKey ? SHORT_LABELS[labelKey] : (h || '(列' + (i + 1) + ')'),
      inList: LIST_COLUMNS.indexOf(h) !== -1,
      readonly: READONLY_COLUMNS.indexOf(h) !== -1,
      type: DATE_COLUMNS.indexOf(h) !== -1 ? 'date' : 'text',
      suggestions: suggestions,
    };
  });
}

/** 文字列として保存すべき列か */
function isTextColumn_(header) {
  return TEXT_COLUMN_KEYWORDS.some(function (k) { return header.indexOf(k) !== -1; });
}

/**
 * 列情報と全データを返す（一覧表示用）
 * Date型はHTML側に渡せないので getDisplayValues で文字列として取得する
 */
function getAllData() {
  const sheet = getSheet_();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];

  let rows = [];
  if (lastRow > HEADER_ROW) {
    rows = sheet
      .getRange(HEADER_ROW + 1, 1, lastRow - HEADER_ROW, lastCol)
      .getDisplayValues()
      .map(function (v, i) { return { row: HEADER_ROW + 1 + i, values: v }; })
      .filter(function (r) { return r.values.some(function (c) { return c !== ''; }); }); // 空行は除外
  }

  return { columns: buildColumns_(headers, rows), rows: rows };
}

/**
 * 指定行の1レコードを取得
 * @param {number} rowNumber シート上の行番号
 */
function getRecord(rowNumber) {
  const sheet = getSheet_();
  validateRow_(sheet, rowNumber);
  const lastCol = sheet.getLastColumn();
  const range = sheet.getRange(rowNumber, 1, 1, lastCol);
  return {
    row: rowNumber,
    values: range.getDisplayValues()[0],
    // 数式が入っているセルは編集不可にする（上書きで数式が消えるのを防ぐ）
    formulaCols: range.getFormulas()[0].map(function (f) { return f !== ''; }),
  };
}

/**
 * 指定行のレコードを更新
 * @param {number} rowNumber シート上の行番号
 * @param {string[]} newValues 列順の新しい値
 * @param {string[]} originalValues 編集開始時の値（他の人の同時編集を検知するため）
 */
function updateRecord(rowNumber, newValues, originalValues) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000); // 同時保存を防ぐ
  try {
    const sheet = getSheet_();
    validateRow_(sheet, rowNumber);
    const lastCol = sheet.getLastColumn();
    if (newValues.length !== lastCol) throw new Error('列数が一致しません。画面を再読み込みしてください');

    const headers = sheet.getRange(HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
    const range = sheet.getRange(rowNumber, 1, 1, lastCol);
    const current = range.getDisplayValues()[0];
    const formulas = range.getFormulas()[0];

    // 編集中に他の人が同じ行を変更していたら保存しない
    if (current.join('\u0000') !== originalValues.join('\u0000')) {
      throw new Error('編集中に他の人がこの行を変更しました。再読み込みしてやり直してください');
    }

    newValues.forEach(function (v, i) {
      const h = headers[i];
      if (formulas[i] !== '') return;                    // 数式セルはスキップ
      if (READONLY_COLUMNS.indexOf(h) !== -1) return;    // 編集不可の列はスキップ
      if (v === current[i]) return;                      // 変更なしはスキップ

      const cell = sheet.getRange(rowNumber, i + 1);
      if (isTextColumn_(h)) {
        cell.setNumberFormat('@'); // 書式を「書式なしテキスト」にして先頭の0を保持
      }
      cell.setValue(v);
    });

    return getRecord(rowNumber); // 保存後の値を返す
  } finally {
    lock.releaseLock();
  }
}

/** 行番号がデータ範囲内か確認 */
function validateRow_(sheet, rowNumber) {
  if (!Number.isInteger(rowNumber) || rowNumber <= HEADER_ROW || rowNumber > sheet.getLastRow()) {
    throw new Error('不正な行番号です: ' + rowNumber);
  }
}
