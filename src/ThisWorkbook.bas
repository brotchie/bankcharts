VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub test()
    ListCommandBars
    ListCommandBars ("Insert"), True
    'ListCommandBars ("Chart Type"), True
End Sub

Sub ListCommandBars(Optional name As String, Optional listControls As Boolean = False)
    If IsMissing(name) Then
        name = ""
    End If
    
    Dim cmd As CommandBar
    
    If name <> "" Then
        Set cmd = Application.CommandBars(name)
        Debug.Print "== " & cmd.name & " =="
        If listControls Then ListCommandBarControls cmd
    Else
        For Each cmd In Application.CommandBars
            Debug.Print "== " & cmd.name & " =="
            If listControls Then ListCommandBarControls cmd
        Next cmd
    End If
End Sub

Sub ListCommandBarControls(cmd As CommandBar)
    Dim ctl As CommandBarControl
    For Each ctl In cmd.Controls
        Debug.Print ctl.Caption & " " & TypeName(ctl)
    Next ctl
End Sub

Sub OnLoad(ui As IRibbonUI)
    Debug.Print "here"
End Sub

