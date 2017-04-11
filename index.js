// index.js
const XLSX = require("xlsx");
const Utils = XLSX.utils; // XLSX.utilsのalias

// Workbookの読み込み
const book = XLSX.readFile("test.xlsx");

/* === Sheet1の合計 === */
// Sheet1読み込み
const sheet1 = book.Sheets["Sheet1"];
// セルの範囲
const range = sheet1["!ref"]; //B2:B4
// セル範囲を数値表現に変換
const decodeRange = Utils.decode_range(range); //{ s: { c: 1, r: 1 }, e: { c: 1, r: 3 } }

// Sheet1に記載されている数値を合計する
let value = 0;
for (let rowIndex = decodeRange.s.r; rowIndex <= decodeRange.e.r; rowIndex++) {
  for (let colIndex = decodeRange.s.c; colIndex <= decodeRange.e.c; colIndex++) {
    // 数値表現をセルアドレスに変換
    const address = Utils.encode_cell({ r: rowIndex, c:colIndex });
    const cell = sheet1[address];
    if (typeof cell !== "undefined" && typeof cell.v !== "undefined") {
      value += cell.v; //cell: { t: 'n', v: 100, w: '100' }
    }
  }
}
console.log(`合計= ${value}`);

/* === Sheet2のC2に現在日時を書き込む === */
// 現在日時
const d = new Date();
const now = `${d.getFullYear()}/${padLeft(d.getMonth() + 1)}/${padLeft(d.getDate())} ${padLeft(d.getHours())}:${padLeft(d.getMinutes())}:${padLeft(d.getSeconds())}`;
// Sheet2読み込み
const sheet2 = book.Sheets["Sheet2"];
// 現在日時を登録
sheet2["C2"] = { t: "s", v: now, w: now };
// 範囲を更新
sheet2["!ref"] = "B2:C2";
book.Sheets["Sheet2"] = sheet2;

// ファイルを書き込み
XLSX.writeFile(book, "test.xlsx");

/**
 * 文字列を左詰めする
 * 
 * @param {*} value
 * @param {string} [ch] default: "0"
 * @param {string} [len] default: 2
 */
function padLeft(value, ch = "0", len = 2) {
  const str = "" + value; //toString
  const s = ch.repeat(len);
  return (s + str).slice(len * -1);
}
