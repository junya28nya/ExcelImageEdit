# Excel画像処理
Excelのセル上に画像を展開し、画像処理を行うExcel VBA プログラム。

![Excelのセル上に展開された画像](https://github.com/junya28nya/ExcelImageEdit/image/ImageOnExcel.png)

対応画像形式
BMP (MicroSoftペイントでBMP形式で書きだしたもの。他の仕様には対応していない)


## ReadBMP.vb
BMP形式の画像を読み込み、Excelの1セルに1ドット塗りつぶし、画像を展開する。
BMPの読み込みをバイナリからフルスクラッチしているため、MicroSoftペイント以外で書き出したのBMPの読み込みは正常にできるかはわからない。

## ToneMapping
読み込んだBMP画像に色調補正をかける。

## ImageFilter
読み込んだ画像に3*3のフィルターをかける。
