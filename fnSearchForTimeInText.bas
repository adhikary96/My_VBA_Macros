Attribute VB_Name = "fnSearchForTimeInText"
Function searchForTimeInText(InputFilePath As String, Text As String, givenTimeFormat As String) As String
' ' ==========================================================================================================================
' ' Function: SearchForTimeInText
' ' Description: Function searches for a particular Text (expected to be Start time and End time here) in the file given in the FilePath and returns the time
' ' Inputs Required: File Path in Cell B2, Time Format in Cell B3
' ' Comments: Mentioned before each section as required.
' ' Author: Deepraj Adhikary
' ' Date of Creation: Wednesday, 11 September, 2019
' ' ==========================================================================================================================
    Dim fileInput As Integer
    Dim textLine As String
    Dim postext As String
    Dim timeFormatLen As Integer
    Dim timeFound As String
    
    timeFormatLen = Len(givenTimeFormat)
    
    fileInput = FreeFile
    Open InputFilePath For Input As #fileInput
    Do Until EOF(1)
        Line Input #fileInput, textLine
'        Debug.Print textLine
        postext = InStr(textLine, Text)
        timeFound = Mid(textLine, postext + Len(Text) + 2, timeFormatLen)
'        Debug.Print Text & "-> Starts from :" & postext & " -> Length: " & Len(Text)
'        Debug.Print temp
    Loop
    searchForText = timeFound
End Function

