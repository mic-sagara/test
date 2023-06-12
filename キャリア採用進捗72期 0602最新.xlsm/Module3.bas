Attribute VB_Name = "Module3"
Option Explicit

Sub CreateWorkbookandClose()

If MsgBox("エクスポートしますか？", vbYesNo) = vbYes Then
Else: GoTo Continue
End If
Application.ScreenUpdating = False

    Dim L_Row As Long
    Dim L_Column As Long
    Dim U_Name As Variant
    Dim D_Time As Variant
    
    U_Name = Environ("USERNAME")
    Debug.Print U_Name

    L_Row = Cells(3, 10000).End(xlUp).Row
    L_Column = Cells(5, 10000).End(xlToLeft).Column
    
    'Range(Cells(3, 3), Cells(L_Row, L_Column)).Copy
    Cells(5, 3).CurrentRegion.Copy
    
    Workbooks.Add
    
    Cells(2, 2).PasteSpecial
    
    Columns("B:HH").AutoFit
    
    ActiveSheet.Name = "集計結果"
    
    ActiveWorkbook.SaveAs "C:\Users\" & U_Name & "\OneDrive - MIC株式会社\デスクトップ\" & "集計結果" & _
                                    Format(Now, "yyyy-MM-dd-hh-mm") & ".xlsm", _
                                    FileFormat:=xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
    
    ' "C:\Users\" & U_Name & "\Desktop\" & "集計結果" & ".xlsm",
    'C:\Users\j_hyun\OneDrive - MIC株式会社\デスクトップ
'C:\Users\" & U_Name & "\OneDrive - MIC株式会社\デスクトップ


ActiveWorkbook.Close

Application.ScreenUpdating = True

MsgBox "デスクトップにエクスポート完了しました。"

Continue:
End Sub

