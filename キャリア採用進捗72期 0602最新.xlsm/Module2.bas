Attribute VB_Name = "Module2"
Option Explicit
Public Sub �s�폜��s�B��()

'   �s�폜s
    Dim B_Row As Long
    Dim i As Long
    Dim ws3 As Worksheet
    Dim colCount As Long
    Dim Headline_Color As Long
    Dim Hide_Color As Long
    Dim sum As Long
    Dim DarkBlue As Long
    
    Set ws3 = ThisWorkbook.Worksheets("�T�����ʁi�S�̍ŐV�j")
    
    Headline_Color = RGB(197, 220, 255)
    Hide_Color = RGB(232, 241, 255)
    DarkBlue = RGB(0, 32, 96)
    B_Row = ws3.Cells(6, 2).End(xlDown).Row
    colCount = ws3.Cells(6, 2).End(xlToRight).Column

    For i = B_Row To 7 Step -1
        sum = Application.WorksheetFunction.sum(Range(Cells(i, 4), Cells(i, colCount)))
        If ws3.Cells(i, 2).Interior.Color = Hide_Color And sum = 0 Then
            ws3.Rows(i).Delete
        ElseIf ws3.Cells(i, 2).Interior.Color = Headline_Color Then
            If sum <> 0 Then
           
                ws3.Cells(i, 1) = "+"
            Else
               
                ws3.Cells(i, 1) = ""
           End If
        End If
    Next i
    
    B_Row = ws3.Cells(6, 2).End(xlDown).Row
    ws3.Cells(B_Row + 1, 2) = "Data"
    Debug.Print B_Row
    
'    �s�S�\��
    ws3.Range("B7", Cells(Rows.Count, 2).End(xlUp)).EntireRow.Hidden = False
            
'    �s�B��
    For i = B_Row To 7 Step -1
        If ws3.Cells(i, 2).Interior.Color = Hide_Color Then
            ws3.Rows(i).Hidden = True
        End If
    Next i
End Sub

Sub �N���A()
    Dim B_Row As Long
    Dim i As Long
    Dim ws3 As Worksheet
    
    Set ws3 = ThisWorkbook.Worksheets("�T�����ʁi�S�̍ŐV�j")
    
    ws3.Range("B7", Cells(Rows.Count, 2).End(xlUp)).EntireRow.Hidden = False
    
End Sub


