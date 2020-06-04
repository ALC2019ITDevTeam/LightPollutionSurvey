VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
'光害調査プログラムコード
'2000 年2月27日作成
'2000 年7月16日修正

Dim i, j1, j2 As Integer
Dim x1, x2, y1, y2 As Double

    For i = 4 To 6
       Sheet2.Cells(3, i) = Sheet1.Cells(3, i)
    Next i

    For j = 7 To 70

       Sheet2.Cells(j, 4) = Sheet5.Cells(j - 4, 6)
       Sheet2.Cells(j, 5) = Sheet5.Cells(j - 4, 10)
       Sheet2.Cells(j, 6) = Sheet5.Cells(j - 4, 11)
       Sheet2.Cells(j, 7) = Sheet5.Cells(j - 4, 12)
       Sheet2.Cells(j, 8) = Sheet1.Cells(j, 5)
       Sheet2.Cells(j, 9) = Sheet1.Cells(j, 7)

       x1 = Sheet1.Cells(j, 5)
       If x1 <> 0 Then
         y1 = 19.266
         y1 = y1 - 0.57119 * x1
         y1 = y1 + 0.012264 * x1 ^ 2
         y1 = y1 - 0.00012328 * x1 ^ 3
         y1 = y1 + 0.0000006712 * x1 ^ 4
         y1 = y1 - 0.0000000018974 * x1 ^ 5
         y1 = y1 + 2.1892E-12 * x1 ^ 6
         Sheet2.Cells(j, 10) = y1
       Else
         Sheet2.Cells(j, 10) = "    /"
       End If

       x2 = Sheet1.Cells(j, 7)
       If x2 <> 0 Then
         y2 = 19.266
         y2 = y2 - 0.57119 * x2
         y2 = y2 + 0.012264 * x2 ^ 2
         y2 = y2 - 0.00012328 * x2 ^ 3
         y2 = y2 + 0.0000006712 * x2 ^ 4
         y2 = y2 - 0.0000000018974 * x2 ^ 5
         y2 = y2 + 2.1892E-12 * x2 ^ 6
         Sheet2.Cells(j, 11) = y2
       Else
         Sheet2.Cells(j, 11) = "    /"
       End If

    Next j


    j1 = 2
    j2 = 2
    For i = 7 To 70

       If Sheet2.Cells(i, 10) <> "    /" Then
         Sheet3.Cells(j1, 7) = Sheet2.Cells(i, 3)
         Sheet3.Cells(j1, 1) = Sheet2.Cells(i, 4)
         Sheet3.Cells(j1, 2) = Sheet2.Cells(i, 5)
         Sheet3.Cells(j1, 3) = Sheet2.Cells(i, 6)
         Sheet3.Cells(j1, 4) = Sheet2.Cells(i, 7)
         Sheet3.Cells(j1, 5) = Sheet2.Cells(i, 10)
         j1 = j1 + 1
       End If

       If Sheet2.Cells(i, 11) <> "    /" Then
         Sheet4.Cells(j2, 7) = Sheet2.Cells(i, 3)
         Sheet4.Cells(j2, 1) = Sheet2.Cells(i, 4)
         Sheet4.Cells(j2, 2) = Sheet2.Cells(i, 5)
         Sheet4.Cells(j2, 3) = Sheet2.Cells(i, 6)
         Sheet4.Cells(j2, 4) = Sheet2.Cells(i, 7)
         Sheet4.Cells(j2, 5) = Sheet2.Cells(i, 11)
         j2 = j2 + 1
       End If

    Next i


End Sub
