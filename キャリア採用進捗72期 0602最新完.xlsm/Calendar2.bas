Attribute VB_Name = "Calendar2"
Option Explicit


Public Sub カレンダーから入力_Click()
    Dim posData As Object
    Set posData = CreateObject("Scripting.Dictionary")
    
    Set posData = ShowUserFormAtMousePosition()

    
    CalendarForm.StartUpPosition = 0
    CalendarForm.Top = posData("y")
    CalendarForm.Left = posData("x")
   'CalendarForm.Top = 300
    'CalendarForm.Left = -1400
    
    Call CalendarForm.Show
End Sub




