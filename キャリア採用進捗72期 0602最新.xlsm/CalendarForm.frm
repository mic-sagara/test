VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CalendarForm 
   Caption         =   "日付を選択してセルに入力"
   ClientHeight    =   3600
   ClientLeft      =   36
   ClientTop       =   168
   ClientWidth     =   3060
   OleObjectBlob   =   "CalendarForm.frx":0000
End
Attribute VB_Name = "CalendarForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CalendarParts(1 To 42)  As CalendarControl
Private CurrentDate             As Date
 
Private Const GRAY      As Long = -2147483633
Private Const LIGHTBLUE As Long = 16763070
 
Private Sub UserForm_Initialize()
 
    Dim i As Long
    Dim pos As Long
    pos = Range("C4")
    '手動指定の場合は処理を分岐します
    If pos = 0 Then
        StartUpPosition = 0  '手動指定
        Me.Left = Range("B2").Value
        Me.Top = Range("B3").Value
    Else
        StartUpPosition = pos  '表示位置指定
    End If
    
    For i = LBound(CalendarParts) To UBound(CalendarParts)
        Set CalendarParts(i) = New CalendarControl
        Call CalendarParts(i).Bind(Me.Controls("Label" & i))
    Next i
 
    CurrentDate = Date
    Call CreateDays
 
End Sub
 
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Range("B2") = Me.Left
    Range("B3") = Me.Top
End Sub

Private Sub UserForm_Terminate()
 
    Erase CalendarParts
 
End Sub
 
Private Sub TXT日付_Change()
 
    If IsDate(Me.TXT日付.Value) Then
        CurrentDate = Me.TXT日付.Value
        Call CreateDays
    End If
 
End Sub
 
Private Sub TXT日付_Exit(ByVal Cancel As MSForms.ReturnBoolean)
 
    If Not IsDate(Me.TXT日付.Value) Then
        Me.TXT日付.Value = CurrentDate
    End If
 
End Sub
 
Private Sub CMD先月_Click()
 
    CurrentDate = DateAdd("m", -1, CurrentDate)
    Me.TXT日付.Value = Format(CurrentDate, "yyyy/mm")
    Call CreateDays
 
End Sub
 
Private Sub CMD翌月_Click()
 
    CurrentDate = DateAdd("m", 1, CurrentDate)
    Me.TXT日付.Value = Format(CurrentDate, "yyyy/mm")
    Call CreateDays
 
End Sub
 
Private Sub CMD今日_Click()
 
    ActiveCell.Value = Date
 
    Call CalendarForm.Hide
 
End Sub
 
Private Sub CreateDays()
 
    Me.TXT日付.Value = Format(CurrentDate, "yyyy/mm")
 
    Dim TargetDate As Date
        TargetDate = Format(CurrentDate, "yyyy/mm") & "/1"
 
    Dim WeekDayCode As Long
        WeekDayCode = 1
 
    Dim Ctrl As Control
    Dim i As Long
    For i = 1 To 42
        Set Ctrl = Me.Controls("Label" & i)
        Ctrl.Caption = ""
        Ctrl.BackColor = GRAY
        If Month(TargetDate) = Month(CurrentDate) _
        And WeekDayCode >= Weekday(TargetDate) Then
            Ctrl.Caption = Day(TargetDate)
            If TargetDate = Date Then
                Ctrl.BackColor = LIGHTBLUE
            End If
            TargetDate = DateAdd("d", 1, TargetDate)
        End If
        WeekDayCode = WeekDayCode + 1
    Next i
 
End Sub
 
Public Sub CopyToActiveCell(ByVal xDate As String)
 
    If xDate = "" Then Exit Sub
 
    ActiveCell.Value = Format(CurrentDate, "yyyy/mm/") & xDate
 
    Call CalendarForm.Hide
 
End Sub

