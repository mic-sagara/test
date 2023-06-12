VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CalendarControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents DateLabel As MSForms.Label
Attribute DateLabel.VB_VarHelpID = -1
 
Public Sub Bind(ByVal Ctrl As MSForms.Control)
    Select Case TypeName(Ctrl)
        Case "Label"
            Set DateLabel = Ctrl
        Case Else
    End Select
End Sub
 
Private Sub DateLabel_Click()
    Call CalendarForm.CopyToActiveCell(DateLabel.Caption)
End Sub
