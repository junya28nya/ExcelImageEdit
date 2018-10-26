Option Explicit On

'BMP�t�@�C���̎d�l
'18-21byte�ɉ��̉�f��
'22-25byte�ɏc�̉�f��
Const WIDTH_POS As Long = 18
Const HEIGHT_POS As Long = 22

Sub ToneMapping()
    Dim openFileName As String      '�J���t�@�C����
    Dim a() As Byte                 'Byte��ǂݍ���
    Dim File_Size As Long           '�ǂݍ��ރt�@�C���̃T�C�Y
    Dim Image_Width_Pixel As Long   '�摜�̉���Pixel��
    Dim Image_Height_Pixel As Long  '�摜�̏c��Pixel��
    Dim Line_Width_Size As Long     '�����C����Byte��
    Dim Line_Last_Size As Long      '�����C���̍Ō�ɂ���ꂽByte��
    Dim Image_Data_Pos As Long      '�C���[�W�f�[�^�̊J�n�ʒu
    Dim image() As Byte             '�摜�̔z��

    Application.ErrorCheckingOptions.BackgroundChecking = False
    ChDir ThisWorkbook.Path & "\"
    openFileName = Application.GetOpenFilename("BMP�摜,*.bmp")
    If openFileName = "False" Then
        MsgBox "BMP�t�@�C����I�����Ă�������"
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

    Call ToneMapping(image(), 0.2, 0.8, 0.1, 0.9)
    'Call ToneMapping2(image())
    Call ChangeColumnWidth(image())
    Call WriteArray2Cells(image())
End Sub

'16To10�ϊ� (Byte�z��, �ŏ���Byte�ʒu, �I����Byte�ʒu)
'   �����o�C�g��16�i����10�i���ɕϊ�����
Function myHex2Dec(a() As Byte, First As Long, Last As Long) As Long
    Dim i As Long
    Dim str As String
    str = ""

    For i = Last To First Step -1
        str = str & Right("00" & Hex(a(i)), 2)
    Next

    myHex2Dec = CInt("&H" & str)
End Function

'BMP�摜��1�s����byte�����v�Z
'BMP�摜��1�s��byte����4�̔{���ɂȂ�悤�A�s���ɉˋ��byte��u���Ă���
Function myCalcLineSize(Width_Pixel As Long) As Long
    Dim Width_Byte As Long
    Width_Byte = Width_Pixel * 3

    If Width_Byte Mod 4 <> 0 Then
        Width_Byte = Width_Byte + (4 - (Width_Byte Mod 4))
    End If

    myCalcLineSize = Width_Byte
End Function

'�摜�̔z����Z���ɓh��
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

'byte��(a())��3�����z��(image(x,y,color))�ɂ����
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

'�z��̒������̍s�Ɨ�̕���2pixel�ɂ���
'height 1pix = 0.75, width 1pix = 0.118
Function ChangeColumnWidth(image() As Byte)
    Dim r As Long
    Dim c As Long
    r = UBound(image, 2) + 1
    c = UBound(image, 1) + 1

    Range(Columns(1), Columns(c)).ColumnWidth = 0.1
    Range(Rows(1), Rows(r)).RowHeight = 1 * 0.75
End Function

'�摜�̐��`�ϊ��I�I
Function ToneMapping(image() As Byte, a1 As Double, a2 As Double, a3 As Double, a4 As Double)
    Dim c As Long, r As Long, color As Long
    Dim rMax As Long, cMax As Long
    rMax = UBound(image, 2)
    cMax = UBound(image, 1)
    Dim image2() As Double
    ReDim image2(cMax, rMax, 2)
    For r = 0 To rMax
        For c = 0 To cMax
            For color = 0 To 2
                image2(c, r, color) = image(c, r, color) / 255
            Next
        Next
    Next
    'Dim a1 As Double, a2 As Double, a3 As Double, a4 As Double
    Dim myScale As Double
    Dim x As Double, xNew As Double
    'a1 = 0.8
    'a2 = 0.2
    'a3 = 0.9
    'a4 = 0
    myScale = (a4 - a3) / (a2 - a1)
    For r = 0 To rMax
        For c = 0 To cMax
            For color = 0 To 2
                x = image(c, r, color) / 255

                If x < a1 Then
                    xNew = a3
                ElseIf a2 < x Then
                    xNew = a4
                Else
                    xNew = myScale * (x - a1) + a3
                End If

                image(c, r, color) = xNew * 255
            Next
        Next
    Next
End Function

'S���̃g�[���J�[�u(���K���z�m�����x�֐��̐ϕ�)
'�ɂ�����`�Z�x�ϊ�
Function ToneMapping3(image() As Byte)
    Dim c As Long, r As Long, color As Long
    Dim x As Byte
    Dim rMax As Long, cMax As Long
    rMax = UBound(image, 2)
    cMax = UBound(image, 1)
    Dim table(255) As Double, tabledx(255) As Double, table255(255) As Byte
    Dim i As Long

    Dim a1 As Double, a2 As Double, a3 As Double, a4 As Double, myScale As Double

    'table��-3����3�Ő��K��
    '�}3SD�̐��K���z��ϕ����邽��
    a1 = 0
    a2 = 255
    a3 = -3
    a4 = 3
    myScale = (a4 - a3) / (a2 - a1)
    For i = 0 To 255
        table(i) = myScale * (i - a1) + a3
    Next

    'table�𐳋K���z�̊m�����x�֐��Ōv�Z
    For i = 0 To 255
        table(i) = (1 / Sqr((2 * 3.1416))) * Exp(-(table(i) ^ 2 / 2))
    Next

    'table��ϕ�����tabledx�ɑ��
    For i = 0 To 255
        If i = 0 Then
            tabledx(i) = 0
        Else
            tabledx(i) = tabledx(i - 1) + table(i)
        End If
    Next

    'tabledx�����f�l�u���p�̃��b�N�A�b�v�e�[�u��table255�����߂�
    a1 = 0
    a2 = tabledx(255)
    a3 = 0
    a4 = 255
    myScale = (a4 - a3) / (a2 - a1)
    For i = 0 To 255
        table255(i) = CByte(myScale * (tabledx(i) - a1) + a3)
    Next

    'image()��table255�ŔZ�x�ϊ�
    For r = 0 To rMax
        For c = 0 To cMax
            For color = 0 To 2
                x = image(c, r, color)
                image(c, r, color) = table255(x)
            Next
        Next
    Next

    'table255�̒l��\��
    'Worksheets.Add
    'For i = 0 To 255
    '    Cells(i + 1, 1) = i
    '    Cells(i + 1, 2) = table255(i)
    'Next
End Function