VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FootballChartUI 
   Caption         =   "Insert Football Field Chart"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   OleObjectBlob   =   "FootballChartUI.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FootballChartUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_refedit As ModelessRefEdit

' Cancel button closes form.
Private Sub btnCancel_Click()
    Me.Hide
End Sub

Private Sub btnInsert_Click()
    Dim generator As New FootballChartGenerator
    Debug.Print txtSource.text
    Debug.Print txtLabels.text
    Debug.Print txtTableDestination.text
    
    generator.GenerateChart Range(txtSource.text), Range(txtLabels.text), Range(txtTableDestination.text)
    
    Me.Hide
End Sub

Private Sub txtLabels_Enter()
    SelectHandlingErrors txtLabels.text
End Sub

Private Sub txtSource_Enter()
    SelectHandlingErrors txtSource.text
End Sub

Private Sub txtTableDestination_Enter()
    SelectHandlingErrors txtTableDestination.text
End Sub

Private Sub UserForm_Initialize()
    Set m_refedit = New ModelessRefEdit
    Set m_refedit.form = Me
    m_refedit.RegisterRefEdit txtSource
    m_refedit.RegisterRefEdit txtLabels
    m_refedit.RegisterRefEdit txtTableDestination
    
    If Application.Selection.Columns.Count = 2 Then
        txtSource.text = Application.Selection.Columns(2).Address
        txtLabels.text = Application.Selection.Columns(1).Address
        txtTableDestination.SetFocus
    End If
End Sub




Private Sub SelectHandlingErrors(target As String)
    On Error Resume Next
    With Application.Range(target)
        If .Worksheet <> ActiveWorkbook.ActiveSheet Then .Worksheet.Activate
        .Select
    End With
    On Error GoTo 0
End Sub


