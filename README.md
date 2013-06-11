excel_vba
=========
excelをDBとして利用できるように、vbaを使ってカスタマイズ

テスト用モジュール : verify.bas
----------
これから作成する関数のテスト用モジュール

シートの操作クラス : clSheet.cls
----------
  * 名前を指定してSheetを作成。同名のシートが存在した場合その中身を削除する。`initSheet(***)`
  * Sheet内のデータ領域をArrayに格納する。`getAllDataAsArray(***) As Boolean`
  * 指定列の最後の行までのデータを取得しArrayに格納。`getColDataAsArray(***) As Boolean`
    
