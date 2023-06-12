Attribute VB_Name = "Module4"
Option Explicit

#If VBA7 Then
    Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
    Declare PtrSafe Function MonitorFromPoint Lib "user32" (ByVal X As Long, ByVal Y As Long, ByVal dwFlags As Long) As LongPtr
    Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
    Declare Function MonitorFromPoint Lib "user32" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
    Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Type POINTAPI
    X As Long
    Y As Long
End Type

Public Function ShowUserFormAtMousePosition() As Object

    Dim cursorPos As POINTAPI
    Dim hMonitor As LongPtr
    Dim UserForm As UserForm1
    Dim xVirtualScreen As Long
    Dim yVirtualScreen As Long
    Dim posData As Object
    
    Set posData = CreateObject("Scripting.Dictionary")
    
    GetCursorPos cursorPos
    
    hMonitor = MonitorFromPoint(cursorPos.X, cursorPos.Y, 2) ' MONITOR_DEFAULTTONEAREST
    
        xVirtualScreen = GetSystemMetrics(76) ' SM_XVIRTUALSCREEN
        yVirtualScreen = GetSystemMetrics(77) ' SM_YVIRTUALSCREEN
        
        
        posData.Add "x", CLng(cursorPos.X * 0.6)
        posData.Add "y", CLng(cursorPos.Y * 0.6)
        
        Debug.Print posData("x"), posData("y")
        
        'Debug.Print cursorPos.X & " - " & xVirtualScreen / 2 & " = " & (cursorPos.X + xVirtualScreen / 2) * 0.6
        
        'If xVirtualScreen < 0 Then
        '    posData.Add "x", CLng((cursorPos.X + xVirtualScreen) * 0.6)
        'Else
        '    posData.Add "x", CLng((cursorPos.X - xVirtualScreen) * 0.6)
        'End If
        
        'If yVirtualScreen < 0 Then
        '    posData.Add "y", CLng((cursorPos.Y + yVirtualScreen) * 0.6)
        'Else
        '    posData.Add "y", CLng((cursorPos.Y - yVirtualScreen) * 0.6)
        'End If
        
        'Debug.Print posData("x") & "," & posData("y")
    
    
    Set ShowUserFormAtMousePosition = posData
    
End Function


