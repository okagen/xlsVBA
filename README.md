
配列処理クラス : clArray.cls
=========

1. Convert the contents in Collection object into a two-dimensional array. / Collectionの中身を2次元配列(Array)に変換。`cnvCollToArr`



ファイル処理クラス : clFiles.cls
=========

1. Export some specified modules into the specified directory. / 指定されたモジュールを指定したディレクトリへエクスポートする。`exportModules`

1. Get the full path of all Excel files under the specified directory / 指定したディレクトリ配下にある全Excelファイルのフルパス取得。 `getAllXlsFilePathCol`

1. Copy some specified sheets and modules from one excel book into another. / 指定されたシート、モジュールを、新しいブックの中にコピーする。`copySheetsAndModules`

1. Copy some specified modules from one excel book into another. / 指定されたモジュールを別のブックの中にコピーする。`copyModules`

1. Copy some specified sheets from one excel book into another. / Excelファイル内の指定されたシートを別のExcelファイルコピーする。`copySheets`

1. Copy some sheets with specified prefix characters from one book into another book. / ブック内の指定された接頭文字を持つシートを、別のブックの中にコピーする。`copySheetsWithPrefix`

シート処理クラス : clSheet.cls
=========

1. Convert all cells on a worksheet to values. / 指定したシートの全てのセルを「値」に変換。`convAllCellsOnSheetToValues`

1. Store the specified area in Sheet into Array. Before calling this function, need to check whether the target sheet has existed or not. / Sheet内の指定した領域をArrayに格納する。呼び出し元の関数で、対象シートが存在するかどうかを確認しておく。`getDataAsArray`

1. Create several sheets by specifying their names. If a sheet with the same name exists, delete its contents. / 名前を指定してSheetを作成。 同名のシートが存在した場合その中身を削除する。`initSheet`

1. Create several sheets by specifying their names.If a sheet with the same name exists, add (#) to the end of the sheet name and count up the number. / 名前を指定してSheetを作成。 同名のシートが存在した場合、シート名末尾に(#)を付けてカウントアップ。`newSheet`

1. Copy the specified sheet with a new name. If a sheet with the same name exists, add (#) to the end of the sheet name and count up the number. / 既存Sheetを、指定した名前のシートにコピー。同名のシートが存在した場合、シート名末尾に(#)を付けてカウントアップ。`copySheet`

1. Check Module existence by specifying the name. / 名前を指定してModuleの有無をチェック。`existModule`

1. Check Module existence by specifying the name. / 名前を指定してSheetの有無をチェック。`existSheet`

1. Check Module existence by specifying the name with wild-card character. / 名前を指定(ワイルドカード可)してSheetの有無をチェック。`existSheetWithWildCardCharacter`



複数のシートをまたいだ処理クラス : clSheets.cls
=========

1. Hide the sheets with the specified prefix character. / 指定した接頭文字を持つシートを非表示にする。`hideSheetsWithPrefix`

1. Show the sheets with the specified prefix character. / 指定した接頭文字を持つシートを表示する。`showSheetsWithPrefix`

1. Convert all cells on some worksheets to values. / 指定した複数シートの全てのセルを「値」に変換。`convAllCellsOnSheetsToValues`

1. Delete sheets except specified.  指定したシート以外のシートを削除する。`deleteUnSpecifiedSheets`
 

# ↓↓↓ 過去に作成した下記関数群を現在見直し中。

excel_vba
=========
Using VBA to customize a Excel file so that it can be used as DB.  
excelをDBとして利用できるように、vbaを使ってカスタマイズ


サンプルクラス : sampletool.cls
----------
The sample is made by placing some controls on the Excel sheet and calling some common modules from them.  
Excelのシート上にコントロールを配置して、それらから共通モジュールを呼び出す形でサンプルを作っています。


***

グローバル変数、定数設定モジュール : g_list.bas
----------
Module to initialize some global variables and constants. This module is called when the `Private Sub Workbook_Open ()` event in ThisWorkbook occurred.  
グローバル変数、定数をイニシャライズするモジュールを作っています。このモジュールはThisWorkbook内の`Private Sub Workbook_Open()`イベントが発生した際に呼び出してます。

サンプルモジュール : sample.bas
----------
Sample macro using a class. It's still rough, I will revise it properly when I need it.  
クラスを使って、サンプルマクロを組んでみます。出来は荒いですが、ちゃんと作るときにちゃんとします。  
  * パーツマスターシートを生成。`createPartsMasterSheet`

テスト用モジュール : verify.bas
----------
Function test module. For example,  a test function as `verify_clSheet_getAllDataAsArray ()` will be created for testing a method `getAllDataAsArray (***) As Boolean` in the clSheet class.  
これから作成する関数のテスト用モジュール。例えば`clSheet`クラスの`getAllDataAsArray(***) As Boolean`というメソッドをテストする場合、`verify_clSheet_getAllDataAsArray()`というテスト用の関数を作っています。

***

ActiveXコントロール操作クラス : clAxCtrl.cls
----------
  * ActiveXコントロールのcheckBoxを指定列に複数配置する。その際checkBoxの値と、配置先のセルの値とリンクさせた状態にする。`putChkBoxesV(***) As Boolean`

Array処理クラス : clDatArr.cls
----------
  * 2次元配列(arrA)を、2次元配列(dat)に追加して返す。`addArray(***) As Boolean`
  * 2次元配列(arr)を、(newRow, newCol)の2次元配列に整形して返す。`formatArray(***) As Boolean`
  * 2次元配列(arr)の指定列に、1列挿入し値を埋める。処理後の2次元配列は1列増える。`insertColIntoArray(***) As Boolean`
  * 2次元配列(arr)の指定列を削除。処理後の2次元配列は1列減る。`removeColFromArray(***) As Boolean`
  * 2次元配列で、あるレコード(行)のすべての要素(列)がEmptlyの場合、削除する。`removeEmptyRecord(***) As Boolean`
  * 2次元配列の中に、同じレコード(行)が存在した場合、一つを残して他のレコードを削除する処理を追加する。`removeDuplication(***) As Boolean`
  * 2次元配列を1行ずつCollectionに入れなおす。`cnvArrToColl(***) As Boolean`
  * Collectionの中身を2次元配列に入れなおす。`cnvCollToArr(***) As Boolean`

ExcelのあるシートをDBとして扱うためのクラス : clDB.cls
----------
  * DBシートを作成。無ければ作る、あれば何もしない。`initDB() As Boolean`
  * DBシートにCollectionを使って値を設定。`setDataColl(***) As Boolean`
  * DBシートにArrayを使って値を設定。`setDataArr(***) As Boolean`
  * DBシートから値をCollectionで取得。'getDataColl(***) As Boolean'
  * DBシートから値をArrayで取得。'getDataArr(***) As Boolean'

ディレクトリ処理クラス : clDir.cls
----------
  * 指定ディレクトリにフォルダを作る。同名のフォルダが存在した場合、フォルダ名末尾に(#)を付けてカウントアップ。`createFolder(***) As Boolean`

ファイル処理クラス : clFiles.cls
----------
  * 指定したディレクトリ配下にある全Excelファイルのフルパス取得。`getAllXlsFilePathCol(***) As Boolean`
  * フルパスのCollectionを受け、ファイル名とフォルダ名の2次元Arrayを返す。`getFolderAndFileNameArr(***) As Boolean`
  * フルパスのCollectionを受け、ファイル名とフォルダ名のCollectionを返す。`getFolderAndFileNameColl(***) As Boolean`
  * ファイルのフルパスを受け、ファイル名と保存されているフォルダ名を返す。`getFolderAndFileName(***) As Boolean`
  * ファイル名を受けてworkbookオブジェクトを取得。`getWorkbookObj(***) As Boolean`
  * 指定フォルダ内のファイルを、別のフォルダに新しいファイル名でコピーする。`copyFiles(***) As Boolean`

メール送信処理クラス : clMail.cls
----------
  * アドレス、タイトル、本文を設定しメーラを起動する。（添付ファイルなし）`openMailer(***) As Boolean`
  * アドレス、タイトル、本文、添付ファイルを設定しoutlookを起動する。`openOutlook(***) As Boolean`

シートの操作クラス : clSheet.cls
----------
  * 指定したRange範囲内にある図形を削除する`deleteObjectInRange(***) As Boolean`
  * 名前を指定してSheetの有無をチェック。`existSheet(***) As Boolean`
  * 名前を指定してSheetを作成。同名のシートが存在した場合その中身を削除する。`initSheet(***)`
  * 名前を指定してSheetを作成。同名のシートが存在した場合、シート名末尾に(#)を付けてカウントアップ。`newSheet(***)`
  * Sheet内のデータ領域をArrayに格納する。`getDataAsArray(***) As Boolean`
  * 指定列の最後の行までのデータを取得しArrayに格納。`getColDataAsArray(***) As Boolean`
  * 指定列の最後の行までのデータを削除。`deleteColData(***) As Boolean`
  * 指定列の行数を取得。指定列の一番下(MAX_ROW)から検索して、値がある行をの数を返す。スタート行まで値が無い場合、スタート行を返す。`getLastRow(***) As Boolean`
  * 指定した文字が、指定列に存在した場合、その行を取得しArrayに格納。`getRowDataVLookUp(***) As Boolean`
  * VLOOKUP関数を用いて、別シートまたは別ファイルのデータを参照する。`setDataByVlookup(***) As Boolean`
  * VLOOKUP関数を用いて、別シートまたは別ファイルのデータを参照する（サイレントモード）。`setDataByVlookupSilently(***) As Boolean`
  * 列番号をアルファベットに変換する。`colNo2Txt(***) As Boolean`
  * 2つの列番号をRangeを表すアルファベットに変換する。`colNo2Rng(***) As Boolean`
  * 列αに設定されたIDと関連するIDが列βに存在した場合、列βの値を使って列αを検索し、レコードを取得する。例えば型番管理されている商品の後継型番をたどって、最新の商品型番を見つけるような時に利用`getSuccessorID(***) As Boolean`

シートをまたいだ処理を行うクラス : clSheets.cls
----------
  * 複数シートの中のデータを結合して、Arrayに格納。`combineSheets(***) As Boolean`
  * 無視するシート名を引数で受け、検索対象Sheetの名前Collectionを作成。`getTargetSheets(***) As Boolean`
  * 指定Sheetの指定列にAutoFilterをかける。`setFiltet(***)`
  * VLOOKUP関数を用いて、複数シートを参照してデータを取得する。`getDataFromSheetsByVlookup(***) As Boolean`

フォルダ操作クラス : clFolder.cls
----------
  * フルパスを指定してフォルダを作成。同名のフォルダが存在した場合、フォルダ名末尾に(#)を付けてカウントアップ。`mkFolder(***) As Boolean`
