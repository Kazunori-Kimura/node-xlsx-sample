# [xlsx](https://www.npmjs.com/package/xlsx) サンプル

npm package [xlsx](https://www.npmjs.com/package/xlsx) の簡単なサンプル

---

## 概要

* Node.js で Excelファイルが扱える [xlsx](https://www.npmjs.com/package/xlsx) を紹介します
* Excelファイルの読み書きができます
* PCにExcel入れて無くても大丈夫

---

## そもそも

* なんで VBA 使わずにわざわざ Node.js？
  - VBAだとMacで動かない
  - VBAのエディタ使いづらい
  - コードのバージョン管理しづらい

---

## そこで xlsx

* Node.js で Excelファイルが扱える
* 対応フォーマットがめっちゃ多い https://github.com/SheetJS/js-xlsx#file-formats
* リファレンス見る限り色々できそう

---

## サンプル

---

## workbook, worksheetの読み込み

```js
const XLSX = require("xlsx");
const Utils = XLSX.utils; // XLSX.utilsのalias
// Workbookの読み込み
const book = XLSX.readFile("test.xlsx");
// Sheet1読み込み
const sheet1 = book.Sheets["Sheet1"];
```

---

## セル範囲の取得

```js
// セルの範囲
const range = sheet1["!ref"]; //B2:B4
// セル範囲を数値表現に変換
const decodeRange = Utils.decode_range(range);
console.log(decodeRange);//=>{ s: { c: 1, r: 1 }, e: { c: 1, r: 3 } }
```

---

## セル範囲をくるくる回して値を取得

```js
// Sheet1に記載されている数値を合計する
let value = 0;
for (let rowIndex = decodeRange.s.r; rowIndex <= decodeRange.e.r; rowIndex++) {
  for (let colIndex = decodeRange.s.c; colIndex <= decodeRange.e.c; colIndex++) {
    // 数値表現をセルアドレス ("A1"など) に変換
    const address = Utils.encode_cell({ r: rowIndex, c:colIndex });
    const cell = sheet1[address];
    if (typeof cell !== "undefined" && typeof cell.v !== "undefined") {
      value += cell.v; //cell: { t: 'n', v: 100, w: '100' }
    }
  }
}
console.log(`合計= ${value}`);
```

---

## 書き込み

範囲の更新を忘れずに (30分ハマった)

```js
// Sheet2読み込み
const sheet2 = book.Sheets["Sheet2"];
// セル更新
sheet2["C2"] = { t: "s", v: "hoge", w: "hoge" };
// 範囲を更新
sheet2["!ref"] = "B2:C2";
book.Sheets["Sheet2"] = sheet2;
// ファイルを書き込み
XLSX.writeFile(book, "test.xlsx");
```

---

## サンプルレポジトリ

https://github.com/Kazunori-Kimura/node-xlsx-sample

