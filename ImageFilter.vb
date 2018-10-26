Option Explicit On

'BMPファイルの仕様
'18-21byteに横の画素数
'22-25byteに縦の画素数
Const WIDTH_POS As Long = 18
Const HEIGHT_POS As Long = 22

Sub ReadBMP()
    Dim openFileName As String      '開くファイル名
    Dim a() As Byte                 'Byte列読み込み
    Dim File_Size As Long           '読み込むファイルのサイズ
    Dim Image_Width_Pixel As Long   '画像の横のPixel数
    Dim Image_Height_Pixel As Long  '画像の縦のPixel数
    Dim Line_Width_Size As Long     '横ラインのByte数
    Dim Line_Last_Size As Long      '横ラインの最後につけられたByte数
    Dim Image_Data_Pos As Long      'イメージデータの開始位置
    Dim image() As Byte             '画像の配列

    Application.ErrorCheckingOptions.BackgroundChecking = False
    ChDir ThisWorkbook.Path & "\"
    openFileName = Application.GetOpenFilename("BMP画像,*.bmp")
    If openFileName = "False" Then
        MsgBox "BMPファイルを選択してください"
        Exit Sub
    End If
    Open openFileName For Binary As #1
    File_Size = LOF(1)
    ReDim a(File_Size)
    Get #1, , a()
    Close #1
    
    Image_Width_Pixel = myHex2Dec(a(), WIDTH_POS, WIDTH_POS + 3)
    Image_Height_Pixel = myHex2Dec(a(), HEIGHT_POS, HEIGHT_POS + 3)
    Line_Width_Size = myCalcLineSize(Image_Width_Pixel)
    Line_Last_Size = Line_Width_Size - Image_Width_Pixel * 3
    Image_Data_Pos = myHex2Dec(a(), 10, 13)
    ReDim image(Image_Width_Pixel - 1, Image_Height_Pixel - 1, 3 - 1)

    Call WriteImage2Array(image(), a(), File_Size)

    Call Filter(image())

    Call ChangeColumnWidth(image())
    Call WriteArray2Cells(image())

    Call Histogram2Graph(image())
End Sub


'16To10変換 (Byte配列, 最初のByte位置, 終わりのByte位置)
'   複数バイトの16進数を10進数に変換する
Function myHex2Dec(a() As Byte, First As Long, Last As Long) As Long
    Dim i As Long
    Dim str As String
    str = ""

    For i = Last To First Step -1
        str = str & Right("00" & Hex(a(i)), 2)
    Next

    myHex2Dec = CInt("&H" & str)
End Function

'BMP画像の1行分のbyte数を計算
'BMP画像は1行のbyte数が4の倍数になるよう、行末に架空のbyteを置いてある
Function myCalcLineSize(Width_Pixel As Long) As Long
    Dim Width_Byte As Long
    Width_Byte = Width_Pixel * 3

    If Width_Byte Mod 4 <> 0 Then
        Width_Byte = Width_Byte + (4 - (Width_Byte Mod 4))
    End If

    myCalcLineSize = Width_Byte
End Function

'画像の配列をセルに塗る
Function WriteArray2Cells(a() As Byte)
    Dim r As Long
    Dim c As Long
    Dim color As Long
    Dim rMax As Long
    Dim cMax As Long
    rMax = UBound(a, 2)
    cMax = UBound(a, 1)
    Application.ScreenUpdating = False
    For r = 0 To rMax
        For c = 0 To cMax
            Cells(r + 1, c + 1).Interior.color = RGB(a(c, r, 0), a(c, r, 1), a(c, r, 2))
        Next
    Next
    Application.ScreenUpdating = True
End Function

'byte列(a())を3次元配列(image(x,y,color))にいれる
Function WriteImage2Array(image() As Byte, a() As Byte, fileSize As Long)
    Dim r As Long
    Dim c As Long
    Dim color As Long
    Dim rMax As Long
    Dim cMax As Long
    Dim i As Long
    rMax = UBound(image, 2)
    cMax = UBound(image, 1)
    i = fileSize - 1

    For r = 0 To rMax
        For c = cMax To 0 Step -1
            For color = 0 To 2
                image(c, r, color) = a(i)
                i = i - 1
            Next
        Next
    Next

End Function

'配列の長さ分の行と列の幅を1pixelにする
'height 1pix = 0.75, width 1pix = 0.118
Function ChangeColumnWidth(image() As Byte)
    Dim r As Long
    Dim c As Long
    r = UBound(image, 2) + 1
    c = UBound(image, 1) + 1

    Range(Columns(1), Columns(c)).ColumnWidth = 0.1
    Range(Rows(1), Rows(r)).RowHeight = 1 * 0.75
End Function

'3×3のフィルター
'現在アクティブのsheetのA1～C3値を使ってフィルターする
Function Filter(image() As Byte)

    Dim fil(2, 2) As Double
    Dim x As Long, y As Long, sum As Double, sumFil As Double
    Dim r As Long, c As Long, color As Long
    Dim rMax As Long, cMax As Long
    rMax = UBound(image, 2)
    cMax = UBound(image, 1)
    Dim imageEx() As Byte
    ReDim imageEx(cMax + 2, rMax + 2, 2)
    For r = 0 To 2
        For c = 0 To 2
            fil(c, r) = Cells(r + 1, c + 1)
            sumFil = sumFil + fil(c, r)
        Next
    Next

    'image()の端に1pixel拡張した配列imageEx()を作成
    '(image(0,0,c)の計算をするときにimage(-1,-1,c)を参照してしまうため)
    '一番上と一番下
    For c = 0 To cMax
        For color = 0 To 2
            imageEx(c + 1, 0, color) = image(c, 0, color)
            imageEx(c + 1, rMax + 2, color) = image(c, rMax, color)
        Next
    Next
    '一番右と一番左
    For r = 0 To rMax
        For color = 0 To 2
            imageEx(0, r + 1, color) = image(0, r, color)
            imageEx(cMax + 2, r + 1, color) = image(cMax, r, color)
        Next
    Next
    '四隅
    For color = 0 To 2
        imageEx(0, 0, color) = image(0, 0, color)
        imageEx(cMax + 2, 0, color) = image(cMax, 0, color)
        imageEx(0, rMax + 2, color) = image(0, rMax, color)
        imageEx(cMax + 2, rMax + 2, color) = image(cMax, rMax, color)
    Next
    '中身
    For r = 0 To rMax
        For c = 0 To cMax
            For color = 0 To 2
                imageEx(c + 1, r + 1, color) = image(c, r, color)
            Next
        Next
    Next

    'フィルターの計算
    For r = 1 To rMax + 1
        For c = 1 To cMax + 1
            For color = 0 To 2
                sum = 0
                For x = 0 To 2
                    For y = 0 To 2
                        sum = sum + imageEx(c + x - 1, r + y - 1, color) * fil(x, y)
                    Next
                Next
                If sumFil <> 0 Then
                    sum = sum / sumFil
                Else
                    sum = Abs(sum)
                    If sum >= 255 Then
                        sum = 255
                    End If
                End If
                image(c - 1, r - 1, color) = CByte(sum)
            Next
        Next
    Next

    Worksheets.Add
    Call ChangeColumnWidth(image())
    Call WriteArray2Cells(image())

End Function

Function WriteImage2Jpeg(image() As Byte)
    Dim r As Long
    Dim c As Long
    r = UBound(image, 2) + 1
    c = UBound(image, 1) + 1
    Dim rc As Long

    Range(Cells(1, 1), Cells(r, c)).CopyPicture

    rc = Shell("mspaint", vbNormalFocus)
    Application.Wait Now + TimeValue("00:00:01")
    SendKeys "^v", True
    SendKeys "^s", True

End Function


'画像のヒストグラムを作成
Function Histogram2Graph(image() As Byte)
    Dim histogram(255) As Long
    Dim r As Long, c As Long
    Dim rMax As Long, cMax As Long
    Dim colorAverage As Double
    Dim colorAverageLong As Long
    rMax = UBound(image, 2)
    cMax = UBound(image, 1)
    For r = 0 To rMax
        For c = 0 To cMax
            colorAverage = (CDbl(image(c, r, 0)) + CDbl(image(c, r, 1)) + CDbl(image(c, r, 2))) / 3
            colorAverage = WorksheetFunction.Round(colorAverage, 0)
            colorAverageLong = CLng(colorAverage)
            histogram(colorAverageLong) = histogram(colorAverageLong) + 1
        Next
    Next
    Worksheets.Add
    ActiveSheet.Name = "ヒストグラム"
    For r = 0 To 255
        Cells(r + 1, 1) = r
        Cells(r + 1, 2) = histogram(r)
    Next
    Cells(r + 1, 1).Value = "合計"
    Cells(r + 1, 2).Value = "=sum(B1:B256)"

    With ActiveSheet.Shapes.AddChart.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Range("B1:B256")
        .ChartGroups(1).GapWidth = 0
    End With

End Function



