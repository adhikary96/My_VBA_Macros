VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufTestCase 
   Caption         =   "WealthX : Docupace - Test Case Designing"
   ClientHeight    =   8865
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8535
   OleObjectBlob   =   "ufTestCase.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufTestCase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cellValue As Range
Dim cellRow As Integer
Dim cellCol As Integer


Private Sub btnAddNewTC_Click()
    tbCellValue.Enabled = False
    tbTCName = ""
    tbTCDescription = ""
    tbTCStepNo = ""
    tbExpectedResult = ""
    tbActualResult = ""
    frameTCDetails.Enabled = False
    frameResults.Enabled = False
    btnAddNewTC.Enabled = False
    btnAddNewTC.Enabled = False
End Sub

Private Sub btnClose_Click()
    Unload ufTestCase
End Sub

Private Sub btnSave_Click()
    fillDataInSheet
    btnAddNewTC.Enabled = True
End Sub

Private Sub cbEnableCellEdit_Click()
    If (cbEnableCellEdit.Value = True) Then
        tbCellValue.Enabled = True
    Else
        tbCellValue.Enabled = False
    End If
End Sub

Private Sub cbPrevTC_Click()
    If (cbPrevTC.Value = True) Then
        If (cellValue.Row <> 1) Then
            tbTCName = Cells(cellValue.Row - 1, cellValue.Column)
        Else
            MsgBox "You are at the very first row." & vbNewLine & "There is no previous cell available"
            cbPrevTC = False
        End If
    End If
End Sub

Private Sub btnSelectCell_Click()
    Set cellValue = Application.InputBox(prompt:="Select the cell from where to start entering the Test cases.", Title:="Enter Test Cases", Type:=8)
    tbCellValue = cellValue.Address(RowAbsolute:=False, ColumnAbsolute:=False)
    tbCellValue.Enabled = False
    frameTCDetails.Enabled = True
    frameResults.Enabled = True
End Sub

Private Sub UserForm_Initialize()
    tbCellValue.Enabled = False
    tbTCName = ""
    tbTCDescription = ""
    tbTCStepNo = ""
    tbExpectedResult = ""
    tbActualResult = ""
    frameTCDetails.Enabled = False
    frameResults.Enabled = False
    btnAddNewTC.Enabled = False
    
End Sub

Private Sub fillDataInSheet()
    cellRow = cellValue.Row
    cellCol = cellValue.Column
    Cells(cellRow, cellCol) = tbTCName.Value
    Cells(cellRow, cellCol + 1) = tbTCStepNo.Value
    Cells(cellRow, cellCol + 2) = tbTCDescription.Value
    Cells(cellRow, cellCol + 3) = tbExpectedResult.Value
    Cells(cellRow, cellCol + 4) = tbActualResult.Value
End Sub

