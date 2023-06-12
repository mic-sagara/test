Attribute VB_Name = "Module1"
Option Explicit


Sub DataInObject()
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim rowCnt As Long
    Dim colCnt As Long
    Dim mainColor As Long
    Dim courseColor As Long
    Dim borderCourseColor As Long
    Dim borderMainColor As Long
    
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    
    Dim dateFrom As Date
    Dim dateTo As Date
    Dim dateRange As Long
    Dim dateValue As String
    Dim searchCourse As String
    
    Dim jobNames As Variant
    Dim courses As Variant
    Dim jobName As String
    Dim course As String
    Dim jobMap As Object
    Dim DateCounts As Object
    Dim JobCounts As Object
    
    Dim mapName As String
    Dim vbAnswer As String
    
    
    Set jobMap = CreateObject("Scripting.Dictionary")
    Set DateCounts = CreateObject("Scripting.Dictionary")
    Set JobCounts = CreateObject("Scripting.Dictionary")
    
    
    Set ws1 = ThisWorkbook.Worksheets("�i���\")
    Set ws2 = ThisWorkbook.Worksheets("�T�����ʁi�S�̍ŐV�j")
    
    'Cell Color
    mainColor = RGB(197, 220, 255)
    courseColor = RGB(232, 241, 255)
    borderMainColor = RGB(20, 40, 119)
    borderCourseColor = RGB(211, 217, 244)
    
    Application.ScreenUpdating = False
    
    rowCnt = ws1.Cells(2, 1).End(xlDown).Row
    
    dateFrom = ws2.Range("E3")
    dateTo = ws2.Range("F3")
    searchCourse = ws2.Range("G3")
    
    dateRange = dateTo - dateFrom
    If dateFrom = 0 Then
        MsgBox ("From����͂��Ă��������B")
        Exit Sub
    End If
    If dateTo = 0 Then
        MsgBox ("To����͂��Ă��������B")
        Exit Sub
    End If
    If searchCourse = "" Then
        MsgBox ("����o�H����͂��Ă�������")
        Exit Sub
    End If
    
    
    'Get Job Name List
    jobNames = GetJobName()
    courses = GetCourseName()
    
    
    For i = 2 To rowCnt
        'Date Value
         dateValue = ws1.Cells(i, 10)
         
        'DateFrom <= DateValue <= dateTo\
        
        
        '���t���Ȃ��Ƃ��@�G���[�`�F�b�N��
        If dateValue = "" Or ws1.Cells(i, 7) = 0 Or ws1.Cells(i, 8) = 0 Then
            vbAnswer = MsgBox("[ " & i & "��ɏ�񂪂���܂���B�������܂����H" + " ]" & vbCrLf & vbCrLf & _
                              "'������'�������ƕҏW���܂��B" _
                    , vbYesNo)
            
            If vbAnswer = vbNo Then
                Worksheets("�i���\").Activate
                ws1.Rows(i).Activate
                Exit Sub
            Else
                GoTo Continue
            End If
        End If
        
        If CDate(dateValue) >= CDate(dateFrom) And CDate(dateValue) <= CDate(dateTo) Then
            'Not SearchCourse = Continue
            If ws1.Cells(i, 9) <> searchCourse And searchCourse <> "�S��" Then
                GoTo Continue
            End If
            
            mapName = jobName + "_count_" + dateValue
            jobName = ws1.Cells(i, 7)
            course = "_" + ws1.Cells(i, 8)
           
            'Job�@Count Check  Name_Count
            mapName = jobName + "_count_" + dateValue
            If jobMap.Exists(mapName) Then
                jobMap(mapName) = jobMap(mapName) + 1
            Else
                jobMap.Add mapName, 1
            End If
        
            'Course_Count
            mapName = jobName + course + "_" + dateValue
            If jobMap.Exists(mapName) Then
                jobMap(mapName) = jobMap(mapName) + 1
            Else
                jobMap.Add mapName, 1
            End If
        End If
Continue:
    Next i
    
    'Printing Setting
    rowCnt = 7
    colCnt = 2
    

    
    'Clear Data
    ws2.Rows("6").ClearContents
    ws2.Rows("7:10000").Clear
    ws2.Cells(6, 1) = "+"
    ws2.Cells(6, 2) = "�E��"
    ws2.Cells(6, 3) = " ���v "
    
    
    'B Column Print
    For i = 0 To UBound(jobNames)
        ws2.Cells(rowCnt, colCnt) = jobNames(i)
        ws2.Rows(rowCnt).Interior.Color = mainColor
        ws2.Cells(rowCnt, colCnt).Font.Bold = True
    
        For j = 0 To UBound(courses)
            rowCnt = rowCnt + 1
            ws2.Cells(rowCnt, colCnt) = courses(j)
            ws2.Rows(rowCnt).Interior.Color = courseColor
        Next j
        
        rowCnt = rowCnt + 1
    Next i
    
    colCnt = colCnt + 1
    
    'Date Print
    For i = 0 To dateRange
        rowCnt = 7
        colCnt = colCnt + 1
        ws2.Cells(rowCnt - 1, colCnt) = CStr(dateFrom + i)
            
        For j = 0 To UBound(jobNames)
            'Main Count Print
            ws2.Cells(rowCnt, colCnt) = jobMap(jobNames(j) + "_count_" + CStr(dateFrom + i))
            
            '���v�@Count
            If DateCounts.Exists(CStr(dateFrom + i)) Then
                DateCounts(CStr(dateFrom + i)) = DateCounts(CStr(dateFrom + i)) + jobMap(jobNames(j) + "_count_" + CStr(dateFrom + i))
            Else
                DateCounts.Add CStr(dateFrom + i), jobMap(jobNames(j) + "_count_" + CStr(dateFrom + i))
            End If
            
            rowCnt = rowCnt + 1
            
            If jobMap(jobNames(j) + "_count_" + CStr(dateFrom + i)) <> "" Then
                'Main Count������ꍇ�o��
                For k = 0 To UBound(courses)
                    'Course Count Input
                    ws2.Cells(rowCnt, colCnt) = jobMap(jobNames(j) + "_" + courses(k) + "_" + CStr(dateFrom + i))
                    rowCnt = rowCnt + 1
                Next k
            Else
                'Main Count���Ȃ��ꍇ
                rowCnt = rowCnt + UBound(courses) + 1
            End If
        Next j
    Next i
    
    Call �s�폜��s�B��
    
    ws2.Columns("B:HH").AutoFit
    ws2.Columns("C:HH").HorizontalAlignment = xlCenter
    ws2.Columns("A").HorizontalAlignment = xlCenter
    
    '�ŏI�s�@�`�F�b�N
    rowCnt = ws2.Cells(6, 2).End(xlDown).Row
    
    'Border Printing & ���v�v�Z
    For i = 7 To rowCnt
         ws2.Cells(i, 3) = Application.WorksheetFunction.sum(Range(Cells(i, 4), Cells(i, 4 + dateRange)))
         ws2.Cells(i, 3).Font.Size = 12
        ws2.Cells(i, 3).Font.Bold = True
           
         
        If ws2.Cells(i, 2).Interior.Color = mainColor Then
                '���v
                ws2.Cells(i, 3).Interior.Color = RGB(0, 32, 96)
                ws2.Cells(i, 3).Font.Color = RGB(255, 255, 255)
                '�f�U�C��
                ws2.Rows(i).Borders(xlEdgeTop).Weight = xlThin
                ws2.Rows(i).Borders(xlEdgeTop).Color = borderMainColor
                
            Else
                ws2.Cells(i, 3).Interior.Color = RGB(207, 218, 247)
                'ws2.Cells(i, 3).Font.Color
                
                ws2.Rows(i).Borders(xlEdgeBottom).Weight = xlThin
                ws2.Rows(i).Borders(xlEdgeBottom).Color = borderCourseColor
        End If
    Next i
    
    '�ŏI�s�`�F�b�N
    i = rowCnt
    Do While ws2.Cells(i, 2) <> 0
        i = i + 1
    Loop
    rowCnt = i - 1
    
    '���v Printing
    ws2.Rows(rowCnt).Interior.Color = RGB(0, 32, 96)
    ws2.Rows(rowCnt).Font.Color = RGB(255, 255, 255)
    ws2.Rows(rowCnt).Font.Bold = True
    ws2.Rows(rowCnt).Font.Size = 12
    ws2.Rows(rowCnt).HorizontalAlignment = xlCenter
    ws2.Cells(rowCnt, 2) = "���v"
    
    'rowCnt,D ~ rowCnt,HH
    ws2.Cells(rowCnt, 3) = "=Sum(D" & CStr(rowCnt) & ":HH" & CStr(rowCnt) + ")"
    
    
        For i = 0 To dateRange
            ws2.Cells(rowCnt, i + 4) = DateCounts(CStr(dateFrom + i))
        Next i
    
    Application.ScreenUpdating = True
    
End Sub
    

Function GetJobName() As String()

    Dim i As Long
    Dim rowCnt As Long
    Dim ws1 As Worksheet
    Dim jobName As String
    Dim jobNames As New Collection
    
    Set ws1 = ThisWorkbook.Worksheets("�i���\")
    
    On Error Resume Next
    rowCnt = ws1.Cells(2, 1).End(xlDown).Row
    
    
    'JobName Check
    For i = 2 To rowCnt
        jobName = ws1.Cells(i, 7).Value
        If jobName <> "" Then
            jobNames.Add jobName, jobName
        End If
    Next i
    
    'Collection To Array
    Dim jobArray() As String: ReDim jobArray(0 To jobNames.Count - 1)
    For i = 1 To jobNames.Count
        jobArray(i - 1) = jobNames.Item(i)
    Next i
    
    GetJobName = jobArray

End Function

Function GetCourseName() As String()
    Dim i As Long
    Dim rowCnt As Long
    Dim ws1 As Worksheet
    Dim course As String
    Dim courses As New Collection
    
    Set ws1 = ThisWorkbook.Worksheets("�i���\")
    On Error Resume Next
    rowCnt = ws1.Cells(2, 1).End(xlDown).Row
    
    'JobName Check
    For i = 2 To rowCnt
        course = ws1.Cells(i, 8).Value
        If course <> "" Then
            courses.Add course, course
        End If
    Next i
    
    
    'Collection To Array
    Dim courseArray() As String: ReDim courseArray(0 To courses.Count - 1)
    For i = 1 To courses.Count
        courseArray(i - 1) = courses.Item(i)
    Next i
    
    
    GetCourseName = courseArray
    

End Function

Sub MoveSheet()
    Dim ws1 As Worksheet
    Set ws1 = ThisWorkbook.Worksheets("�i���\")
    ws1.Activate
    
End Sub










