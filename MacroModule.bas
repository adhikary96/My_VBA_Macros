Attribute VB_Name = "MacroModule"
Sub Extract()
' ' ========================================================================================================
' '         Program: Extract
' '     Description: Macro to extract the data witl multiple enteries in the input sheet to a more ergonomic view in the output sheet
' ' Data Needed: Input data in the 'Input' Sheet in the Columns A & B
' '       Comment: Comments are mentioned before each section for ease of understanding
' '             Author: Deepraj Adhikary
' ' =========================================================================================================
    'Variable Declaration
    Dim book As Workbook
    Dim InputSheet As Worksheet
    Dim OutputSheet As Worksheet
    Dim compareRange As Range
    Dim inRowCount As Integer
    Dim outRowCount As Integer
    Dim compCols As Long
    Dim inRow As Integer
    Dim outRow As Integer
    Dim Data As String
    
    'Setting up the Properties: Workbook, Input and output sheets
    Set book = ThisWorkbook
    Set InputSheet = book.Sheets("Input")
    Set OutputSheet = book.Sheets("Output")
    
    'Getting the numbers of rows in each sheet
    'Input sheet: Overall accounts with matching data
    inRowCount = InputSheet.Range("A" & Rows.Count).End(xlUp).Row
    'Output sheet: unique number of accounts in the list
    outRowCount = OutputSheet.Range("A" & Rows.Count).End(xlUp).Row
    
    'Copy-Paste the accounts list from inouot and remove duplicates
    'Select input and copy accounts list
    InputSheet.Select
    Range("A:A").Copy
    'Select Output sheet and paste data
    OutputSheet.Select
    Range("A:A").Select
    ActiveSheet.Paste
    'Removing duplicates from the output sheet to get list of unique accounts.
    ActiveSheet.Range("A:A").RemoveDuplicates Columns:=Array(1), Header:=xlYes
    
    'Setting up variables for further calculations
    inRow = 2
    outRow = 2
    outCol = 2
    
    'Do while loop 1: to loop through the accounts in the output sheet
    Do While outRow <= outRowCount
        outCol = 2
        inRow = 2
        'Do While Loop 2: to loop through all the acocunts in the input sheet
        Do While inRow <= inRowCount
            cellFind = OutputSheet.Cells(outRow, 1).Text
            'Checking if the account # from the output sheet is present in the input sheet
            'If yes, then get the data corrosponding to the account number from COL 2 and paste it in the adjacent column in the output sheet one after the other
            If (cellFind = InputSheet.Cells(inRow, 1)) Then
                'Data = Data & CStr(InputSheet.Cells(inRow, 2)) & vbNewLine
                OutputSheet.Cells(outRow, outCol) = InputSheet.Cells(inRow, 2)
                outCol = outCol + 1
            End If
            'OutputSheet.Cells(outRow, 2) = Data
            inRow = inRow + 1
        Loop
        outRow = outRow + 1
    Loop
    
End Sub
