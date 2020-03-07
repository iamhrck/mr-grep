const fs = require('fs');
const util = require('util');
const readdir = util.promisify(fs.readdir);
const path = require('path');
const XLSX     = require('xlsx');

const TARGET = 'C:\\Udemy';
const excels = [];

const resultMap = new Map();

/**
 * 引数がExcelファイルか判定します。
 * @param {*} file 
 */
const isExcel = file => {
  return /.*\.xlsx$/.test(file);
}
/**
 * 引数がディレクトリか判定します。
 * @param {*} file 
 */
const isDir = path => {
  return !fs.statSync(path).isFile();
}

/**
 * フォルダ名を構築して返却します。
 * @param {*} dir 
 * @param {*} file 
 */
const constructURL = (dir,file) => {
  return `${dir}\\${file}`;
}

/**
 * 引数で与えられたディレクトリの内容を読み取り、返却します。
 * 読み取れなかった場合は空文字を返却します。
 * @param {*} dirname ディレクトリ名
 */
const readDirAsync = async dirname => {
  let inItems;
  // 現在のディレクトリ読み込み処理
  try {
    inItems = readdir(dirname);
  } catch (err) {
    throw err;
  }

  // ディレクトリ内になにも存在しなければ処理終了
  if (!inItems) {
    throw new Error('current directory dose not have items');
  }
  return inItems;
};

/**
 * 検索対象のディレクトリを検索します。
 * @param {*} dirname 対象ディレクトリ
 */
const searchExcel = async(dirname) => {
  // 引数のフォルダ下の情報を取得する
  const items = await readDirAsync(dirname);
  // 戻り値を調査
  for (const item of items) {
    const targetPath = constructURL(dirname, item);
    if(isDir(targetPath)) {
      // もし対象がディレクトリならもう一度フォルダ探索を行う
      await searchExcel(targetPath);
      continue;
    }
    // もし対象がExcelファイル以外だったらループを繰り返す
    if (!isExcel(targetPath)) {
      continue;
    }
    // Excelファイルのリストに格納
    excels.push(targetPath);
    resultMap.set(targetPath, {});
  }
}

const searchWordInExcel = async() => {
  // excelsの中を探索
  for (const file of excels) {
    const book = XLSX.readFile(file)
    const sheets = book.SheetNames;
    // 各シートを探索
    for(const sheet of sheets) {
      const content = book.Sheets[sheet];
      for (const [cell, value] of Object.entries(content)) {
        if(typeof value === 'string') {
          continue;
        }
        if(Object.values(value).includes('あやぼう')) {
          const result = {
            sheet: sheet,
            cell: cell
          };
          console.log(file, cell, sheet);
          resultMap.set(file, result);
        }
      }
    }
  }
}

/**
 * Mainプログラミング
 */
const Main = async () => {
  await searchExcel(TARGET);
  console.log(excels);
  // Excel内で検索文字列を検索
  await searchWordInExcel();
  console.log(resultMap);
};

Main();
