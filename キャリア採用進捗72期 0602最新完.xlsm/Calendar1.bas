Attribute VB_Name = "Calendar1"
Option Explicit

Private Sub CreateCalendarForm()
    Dim myForm As Object
    Set myForm = ThisWorkbook.VBProject.VBComponents.Add(ComponentType:=3)  '' vbext_ct_MSForm
    With myForm
        .Name = "CalendarForm"
        .Properties("Height") = 310
        .Properties("Width") = 310
        .Properties("Caption") = "日付を選択してセルに入力"
    End With
    Dim myFormDesign As Object
    Set myFormDesign = myForm.Designer
    With myFormDesign.Controls.Add("Forms.TextBox.1")
        .Name = "TXT日付"
        .Width = 144
        .Height = 24
        .Top = 6
        .Left = 78
        .BackColor = 16777215
        .BackStyle = 1
        .ForeColor = 0
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 0
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .MaxLength = 10
        .IMEMode = 3
    End With
    With myFormDesign.Controls.Add("Forms.CommandButton.1")
        .Name = "CMD先月"
        .Width = 30
        .Height = 24
        .Top = 6
        .Left = 42
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 1
        .Caption = "<"
    End With
    With myFormDesign.Controls.Add("Forms.CommandButton.1")
        .Name = "CMD翌月"
        .Width = 30
        .Height = 24
        .Top = 6
        .Left = 228
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 2
        .Caption = ">"
    End With
    With myFormDesign.Controls.Add("Forms.CommandButton.1")
        .Name = "CMD今日"
        .Width = 48
        .Height = 24
        .Top = 252
        .Left = 126
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 10.2
        .Font.Name = "MS UI Gothic"
        .TabIndex = 3
        .Caption = "今日"
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label43"
        .Width = 36
        .Height = 20
        .Top = 40
        .Left = 24
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 255
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 4
        .BorderColor = -2147483633
        .BorderStyle = 1
        .SpecialEffect = 0
        .TextAlign = 2
        .Caption = "日"
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label44"
        .Width = 36
        .Height = 20
        .Top = 40
        .Left = 60
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 5
        .BorderColor = -2147483633
        .BorderStyle = 1
        .SpecialEffect = 0
        .TextAlign = 2
        .Caption = "月"
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label45"
        .Width = 36
        .Height = 20
        .Top = 40
        .Left = 96
        .BackColor = -2147483633
        .BackStyle = 0
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 6
        .BorderColor = -2147483633
        .BorderStyle = 1
        .SpecialEffect = 0
        .TextAlign = 2
        .Caption = "火"
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label46"
        .Width = 36
        .Height = 20
        .Top = 40
        .Left = 132
        .BackColor = -2147483633
        .BackStyle = 0
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 7
        .BorderColor = -2147483633
        .BorderStyle = 1
        .SpecialEffect = 0
        .TextAlign = 2
        .Caption = "水"
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label47"
        .Width = 36
        .Height = 20
        .Top = 40
        .Left = 168
        .BackColor = -2147483633
        .BackStyle = 0
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 8
        .BorderColor = -2147483633
        .BorderStyle = 1
        .SpecialEffect = 0
        .TextAlign = 2
        .Caption = "木"
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label48"
        .Width = 36
        .Height = 20
        .Top = 40
        .Left = 204
        .BackColor = -2147483633
        .BackStyle = 0
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 9
        .BorderColor = -2147483633
        .BorderStyle = 1
        .SpecialEffect = 0
        .TextAlign = 2
        .Caption = "金"
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label49"
        .Width = 36
        .Height = 20
        .Top = 40
        .Left = 240
        .BackColor = -2147483633
        .BackStyle = 0
        .ForeColor = 16711680
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 10
        .BorderColor = -2147483633
        .BorderStyle = 1
        .SpecialEffect = 0
        .TextAlign = 2
        .Caption = "土"
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label1"
        .Width = 36
        .Height = 30
        .Top = 66
        .Left = 24
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 255
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 13
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label2"
        .Width = 36
        .Height = 30
        .Top = 66
        .Left = 60
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 14
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label3"
        .Width = 36
        .Height = 30
        .Top = 66
        .Left = 96
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 15
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label4"
        .Width = 36
        .Height = 30
        .Top = 66
        .Left = 132
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 16
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label5"
        .Width = 36
        .Height = 30
        .Top = 66
        .Left = 168
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 17
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label6"
        .Width = 36
        .Height = 30
        .Top = 66
        .Left = 204
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 18
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label7"
        .Width = 36
        .Height = 30
        .Top = 66
        .Left = 240
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 16711680
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 19
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label8"
        .Width = 36
        .Height = 30
        .Top = 96
        .Left = 24
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 255
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 20
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label9"
        .Width = 36
        .Height = 30
        .Top = 96
        .Left = 60
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 21
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label10"
        .Width = 36
        .Height = 30
        .Top = 96
        .Left = 96
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 22
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label11"
        .Width = 36
        .Height = 30
        .Top = 96
        .Left = 132
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 23
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label12"
        .Width = 36
        .Height = 30
        .Top = 96
        .Left = 168
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 24
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label13"
        .Width = 36
        .Height = 30
        .Top = 96
        .Left = 204
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 25
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label14"
        .Width = 36
        .Height = 30
        .Top = 96
        .Left = 240
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 16711680
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 26
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label15"
        .Width = 36
        .Height = 30
        .Top = 126
        .Left = 24
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 255
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 27
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label16"
        .Width = 36
        .Height = 30
        .Top = 126
        .Left = 60
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 28
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label17"
        .Width = 36
        .Height = 30
        .Top = 126
        .Left = 96
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 29
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label18"
        .Width = 36
        .Height = 30
        .Top = 126
        .Left = 132
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 30
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label19"
        .Width = 36
        .Height = 30
        .Top = 126
        .Left = 168
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 31
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label20"
        .Width = 36
        .Height = 30
        .Top = 126
        .Left = 204
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 32
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label21"
        .Width = 36
        .Height = 30
        .Top = 126
        .Left = 240
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 16711680
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 33
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label22"
        .Width = 36
        .Height = 30
        .Top = 156
        .Left = 24
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 255
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 34
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label23"
        .Width = 36
        .Height = 30
        .Top = 156
        .Left = 60
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 35
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label24"
        .Width = 36
        .Height = 30
        .Top = 156
        .Left = 96
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 36
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label25"
        .Width = 36
        .Height = 30
        .Top = 156
        .Left = 132
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 37
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label26"
        .Width = 36
        .Height = 30
        .Top = 156
        .Left = 168
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 38
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label27"
        .Width = 36
        .Height = 30
        .Top = 156
        .Left = 204
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 39
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label28"
        .Width = 36
        .Height = 30
        .Top = 156
        .Left = 240
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 16711680
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 40
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label29"
        .Width = 36
        .Height = 30
        .Top = 186
        .Left = 24
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 255
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 41
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label30"
        .Width = 36
        .Height = 30
        .Top = 186
        .Left = 60
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 42
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label31"
        .Width = 36
        .Height = 30
        .Top = 186
        .Left = 96
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 43
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label32"
        .Width = 36
        .Height = 30
        .Top = 186
        .Left = 132
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 44
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label33"
        .Width = 36
        .Height = 30
        .Top = 186
        .Left = 168
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 45
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label34"
        .Width = 36
        .Height = 30
        .Top = 186
        .Left = 204
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 46
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label35"
        .Width = 36
        .Height = 30
        .Top = 186
        .Left = 240
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 16711680
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 47
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label36"
        .Width = 36
        .Height = 30
        .Top = 216
        .Left = 24
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 255
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 48
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label37"
        .Width = 36
        .Height = 30
        .Top = 216
        .Left = 60
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 49
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label38"
        .Width = 36
        .Height = 30
        .Top = 216
        .Left = 96
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 50
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label39"
        .Width = 36
        .Height = 30
        .Top = 216
        .Left = 132
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 51
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label40"
        .Width = 36
        .Height = 30
        .Top = 216
        .Left = 168
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 52
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label41"
        .Width = 36
        .Height = 30
        .Top = 216
        .Left = 204
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = -2147483630
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 53
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
    With myFormDesign.Controls.Add("Forms.Label.1")
        .Name = "Label42"
        .Width = 36
        .Height = 30
        .Top = 216
        .Left = 240
        .BackColor = -2147483633
        .BackStyle = 1
        .ForeColor = 16711680
        .Font.Size = 15.75
        .Font.Name = "Times New Roman"
        .TabIndex = 54
        .BorderColor = -2147483642
        .BorderStyle = 0
        .SpecialEffect = 3
        .TextAlign = 2
        .Caption = ""
    End With
End Sub
