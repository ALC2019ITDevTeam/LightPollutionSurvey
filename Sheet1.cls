VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
'光害調査プログラムコード
'2000年2月27日作成
'2000年7月16日修正

'2019年5月19日修正
'修正者　2019年度会員 近藤健

Dim j As Integer
Dim x1 As Double, x2 As Double, y1 As Double, y2 As Double
Dim excobj As Workbook, n As Integer
Set excobj = ThisWorkbook
n = excobj.Worksheets(1).Cells(excobj.Worksheets(1).Rows.Count, 3).End(xlUp).Row

Application.ScreenUpdating = False
    For j = 4 To 6
       excobj.Worksheets(2).Cells(3, j).Value = excobj.Worksheets(1).Cells(3, j).Value
    Next j

    For j = 7 To n
       excobj.Worksheets(2).Cells(j, 4).Value = excobj.Worksheets(5).Cells(j - 4, 6).Value
       excobj.Worksheets(2).Cells(j, 5).Value = excobj.Worksheets(5).Cells(j - 4, 10).Value
       excobj.Worksheets(2).Cells(j, 6).Value = excobj.Worksheets(5).Cells(j - 4, 11).Value
       excobj.Worksheets(2).Cells(j, 7).Value = excobj.Worksheets(5).Cells(j - 4, 12).Value
       excobj.Worksheets(2).Cells(j, 8).Value = excobj.Worksheets(1).Cells(j, 5).Value
       excobj.Worksheets(2).Cells(j, 9).Value = excobj.Worksheets(1).Cells(j, 7).Value
       
       x1 = excobj.Worksheets(1).Cells(j, 5).Value
       If x1 <> 0 Then
            y1 = 19.266
            y1 = y1 - 0.57119 * x1
            y1 = y1 + 0.012264 * x1 ^ 2
            y1 = y1 - 0.00012328 * x1 ^ 3
            y1 = y1 + 0.0000006712 * x1 ^ 4
            y1 = y1 - 0.0000000018974 * x1 ^ 5
            y1 = y1 + 2.1892E-12 * x1 ^ 6
            excobj.Worksheets(2).Cells(j, 10).Value = y1
       Else
            excobj.Worksheets(2).Cells(j, 10).Value = "    /"
       End If
       
       x2 = excobj.Worksheets(1).Cells(j, 7).Value
       If x2 <> 0 Then
            y2 = 19.266
            y2 = y2 - 0.57119 * x2
            y2 = y2 + 0.012264 * x2 ^ 2
            y2 = y2 - 0.00012328 * x2 ^ 3
            y2 = y2 + 0.0000006712 * x2 ^ 4
            y2 = y2 - 0.0000000018974 * x2 ^ 5
            y2 = y2 + 2.1892E-12 * x2 ^ 6
            excobj.Worksheets(2).Cells(j, 11).Value = y2
       Else
            excobj.Worksheets(2).Cells(j, 11).Value = "    /"
       End If
       
       If InStrRev(excobj.Worksheets(2).Cells(j, 10).Value, "/") = 0 Then
         excobj.Worksheets(3).Cells(j - 5, 7).Value = excobj.Worksheets(2).Cells(j, 3).Value
         excobj.Worksheets(3).Cells(j - 5, 1).Value = excobj.Worksheets(2).Cells(j, 4).Value
         excobj.Worksheets(3).Cells(j - 5, 2).Value = excobj.Worksheets(2).Cells(j, 5).Value
         excobj.Worksheets(3).Cells(j - 5, 3).Value = excobj.Worksheets(2).Cells(j, 6).Value
         excobj.Worksheets(3).Cells(j - 5, 4).Value = excobj.Worksheets(2).Cells(j, 7).Value
         excobj.Worksheets(3).Cells(j - 5, 5).Value = excobj.Worksheets(2).Cells(j, 10).Value
       End If
    
       If InStrRev(excobj.Worksheets(2).Cells(j, 11).Value, "/") = 0 Then
         excobj.Worksheets(4).Cells(j - 5, 7).Value = excobj.Worksheets(2).Cells(j, 3).Value
         excobj.Worksheets(4).Cells(j - 5, 1).Value = excobj.Worksheets(2).Cells(j, 4).Value
         excobj.Worksheets(4).Cells(j - 5, 2).Value = excobj.Worksheets(2).Cells(j, 5).Value
         excobj.Worksheets(4).Cells(j - 5, 3).Value = excobj.Worksheets(2).Cells(j, 6).Value
         excobj.Worksheets(4).Cells(j - 5, 4).Value = excobj.Worksheets(2).Cells(j, 7).Value
         excobj.Worksheets(4).Cells(j - 5, 5).Value = excobj.Worksheets(2).Cells(j, 11).Value
       End If
    Next j
Application.ScreenUpdating = True

End Sub
