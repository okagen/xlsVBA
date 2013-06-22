excel_vba
=========
excelをDBとして利用できるように、vbaを使ってカスタマイズ


サンプルクラス : sampletool.cls
----------
Excelのシート上にコントロールを配置して、それらから共通モジュールを呼び出す形でサンプルを作っています。

***

グローバル変数、定数設定モジュール : g_list.bas
----------
グローバル変数、定数をイニシャライズするモジュールを作っています。このモジュールはThisWorkbook内の`Private Sub Workbook_Open()`イベントが発生した際に呼び出してます。

サンプルモジュール : sample.bas
----------
クラスを使って、サンプルマクロを組んでみます。出来は荒いですが、ちゃんと作るときにちゃんとします。  
  * パーツマスターシートを生成。`createPartsMasterSheet`

テスト用モジュール : verify.bas
----------
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

シートの操作クラス : clSheet.cls
----------
  * 指定したRange範囲内にある図形を削除する`deleteObjectInRange(***) As Boolean`
  * 名前を指定してSheetの有無をチェック。`existSheet(***) As Boolean`
  * 名前を指定してSheetを作成。同名のシートが存在した場合その中身を削除する。`initSheet(***)`
  * 名前を指定してSheetを作成。同名のシートが存在した場合、シート名末尾に(#)を付けてカウントアップ。`newSheet(***)`
  * Sheet内のデータ領域をArrayに格納する。`getAllDataAsArray(***) As Boolean`
  * 指定列の最後の行までのデータを取得しArrayに格納。`getColDataAsArray(***) As Boolean`
  * 指定列の最後の行までのデータを削除。`deleteColData(***) As Boolean`
  * 指定した文字が、指定列に存在した場合、その行を取得しArrayに格納。`getRowDataVLookUp(***) As Boolean`

シートをまたいだ処理を行うクラス : clSheets.cls
----------
  * 複数シートの中のデータを結合して、Arrayに格納。`combineSheets(***) As Boolean`
  * 無視するシート名を引数で受け、検索対象Sheetの名前Collectionを作成。`getTargetSheets(***) As Boolean`
  * 指定Sheetの指定列にAutoFilterをかける。`setFiltet(***)`



