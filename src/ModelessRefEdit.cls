VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ModelessRefEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_form As UserForm
Private m_refedits As Dictionary
Private WithEvents m_app As Application
Attribute m_app.VB_VarHelpID = -1

Public Property Set form(a_form As UserForm)
    Debug.Assert m_form Is Nothing
    Set m_form = a_form
End Property

Public Sub RegisterRefEdit(control As control)
    Debug.Assert Not m_refedits.Exists(control)
    m_refedits.Add control, control
End Sub

Private Sub UpdateSelection(target As Range)
    If m_refedits.Exists(m_form.ActiveControl) Then
        Dim control As Object
        Set control = m_form.ActiveControl
        With control
            .text = "'" & target.Parent.name & "'!" & target.Address
            .Enabled = False
            .Enabled = True
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.text)
        End With
    End If
End Sub

Private Sub Class_Initialize()
    Set m_form = Nothing
    Set m_refedits = New Dictionary
    Set m_app = Application
End Sub

Private Sub m_app_SheetSelectionChange(ByVal Sh As Object, ByVal target As Range)
    UpdateSelection target
End Sub
