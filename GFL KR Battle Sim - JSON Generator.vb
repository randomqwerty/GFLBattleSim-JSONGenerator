' This workbook uses modifications of the code provided on these two sites:
' (1) https://superuser.com/questions/1249898/saving-excel-sheet-as-json-file
' (2) https://stackoverflow.com/questions/19371990/how-do-i-replace-a-string-in-a-line-of-a-text-file-using-filesystemobject-in-vba

Option Explicit

Sub KRSimJSON()
    Call CreateDollJSON
    Call CreateJSON("Fairy", Range("FairyJSONPath").Value)  ' similar structure as Equip JSON so uses same code with diff parameter
    Call CreateJSON("Equip", Range("EquipJSONPath").Value)  ' similar structure as Fairy JSON so uses same code with diff parameter
    Call MissionJSON
    Call SimJSON
    MsgBox "Process completed."
End Sub

Public Sub CreateDollJSON()
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim fullFilePath As String: fullFilePath = Range("DollJSONPath").Value
    Dim sheetName As String: sheetName = "Doll"

    Dim fileStream As Object
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 2 'Specify stream type - we want To save text/string data.
    fileStream.Charset = "utf-8" 'Specify charset For the source text data.
    fileStream.Open 'Open the stream And write binary data To the object

    Dim wkb As Workbook
    Set wkb = ThisWorkbook

    Dim wks As Worksheet
    Set wks = wkb.Sheets(sheetName)

    Dim lcolumn As Long: lcolumn = wks.Cells(1, Columns.Count).End(xlToLeft).Column
    Dim lrow As Long: lrow = wks.Cells(Rows.Count, "A").End(xlUp).row
        
    Dim titles() As String
    ReDim titles(lcolumn)
    
    Dim dq As String: dq = """"
    Dim escapedDq As String: escapedDq = "\"""
    Dim twospace As String: twospace = "  "
    Dim fourspace As String: fourspace = "    "
    
    ' Define array of column titles
    Dim i As Integer
    For i = 1 To lcolumn
        titles(i) = wks.Cells(1, i)
    Next i
    
    ' First line of JSON
    fileStream.WriteText "[" & vbNewLine
    
    Dim j As Integer
    Dim cellvalue As String
    
    ' Loop through rows and columns
    For j = 2 To lrow
        ' Only executes if the row has a non-blank ID
        If wks.Cells(j, 1).Value <> "" Then
            
            For i = 1 To lcolumn
                If i = 1 Then
                    fileStream.WriteText twospace & "{" & vbNewLine
                End If
                
                cellvalue = Replace(wks.Cells(j, i), dq, escapedDq)
                fileStream.WriteText fourspace & dq & titles(i) & dq & ": " & dq & cellvalue & dq
                
                If i <> lcolumn Then
                    fileStream.WriteText ","
                End If
                
                fileStream.WriteText vbNewLine
            Next i
            
            fileStream.WriteText twospace & "}"
            
            If j <> lrow Then
                fileStream.WriteText ","
            End If
            
            fileStream.WriteText vbNewLine
        End If
    Next j
    
    fileStream.WriteText "]"
    fileStream.SaveToFile fullFilePath, 2 'Save binary data To disk
End Sub

Public Sub CreateJSON(sheetName As String, fullFilePath As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    Dim fileStream As Object
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 2 'Specify stream type - we want To save text/string data.
    fileStream.Charset = "utf-8" 'Specify charset For the source text data.
    fileStream.Open 'Open the stream And write binary data To the object

    Dim wkb As Workbook: Set wkb = ThisWorkbook
    Dim wks As Worksheet: Set wks = wkb.Sheets(sheetName)

    Dim lcolumn As Long: lcolumn = wks.Cells(1, Columns.Count).End(xlToLeft).Column
    Dim lrow As Long: lrow = wks.Cells(Rows.Count, "A").End(xlUp).row
    
    Dim titles() As String
    ReDim titles(lcolumn)
    
    Dim dq As String: dq = """"
    Dim escapedDq As String: escapedDq = "\"""
    Dim twospace As String: twospace = "  "
    Dim fourspace As String: fourspace = "    "
    
    ' Define array of column titles
    Dim i As Integer
    For i = 1 To lcolumn
        titles(i) = wks.Cells(1, i)
    Next i
    
    fileStream.WriteText "{" & vbNewLine
    
    ' Loop through rows and columns
    Dim j As Integer
    Dim cellvalue As String
    For j = 2 To lrow
        If wks.Cells(j, 1).Value <> "" Then

            fileStream.WriteText twospace & dq & j - 1 & dq & ": {" & vbNewLine
            
            For i = 1 To lcolumn
                
                cellvalue = Replace(wks.Cells(j, i), dq, escapedDq)
                
                fileStream.WriteText fourspace & dq & titles(i) & dq & ": " & dq & cellvalue & dq
                
                If i <> lcolumn Then
                    fileStream.WriteText ","
                End If
                
                fileStream.WriteText vbNewLine
            Next i
            
            fileStream.WriteText twospace & "}"
            
            If j <> lrow Then
                fileStream.WriteText ","
            End If
            
            fileStream.WriteText vbNewLine
        End If
    Next j
    
    fileStream.WriteText "}"
    fileStream.SaveToFile fullFilePath, 2 'Save binary data To disk
End Sub

Sub MissionJSON()
    ' To keep things simple, this simply replaces certain lines in the JSON instead of creating the JSON from scratch

    ' Declare variables
    Const ForReading = 1    '
    Dim fileToRead As String: fileToRead = Range("MissionJSON").Value   ' the path of the file to read
    Dim fileToWrite As String: fileToWrite = fileToRead                 ' the path of a new file (set to be the same as the read file)
    Dim FSO As Object
    Dim readFile As Object      'the file you will READ
    Dim writeFile As Object     'the file you will CREATE (set to be the same as the read file)
    Dim repLine As Variant      'the array of lines you will WRITE
    Dim lrow As Long
    Dim l As Long
    Dim wks As Worksheet
    
    ' Read entire file into an array & close it
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set readFile = FSO.OpenTextFile(fileToRead, ForReading, False)
    repLine = Split(readFile.ReadAll, vbNewLine)
    readFile.Close
    
    Set writeFile = FSO.CreateTextFile(fileToRead, True, False)
    
    Dim dq As String: dq = """"
    Dim escapedDq As String: escapedDq = "\"""
    Dim twospace As String: twospace = "  "
    Dim fourspace As String: fourspace = "    "
    
    ' Set worksheet name and number of rows to loop through
    Set wks = ThisWorkbook.Sheets("Mission_Act_Info")
    lrow = wks.Cells(Rows.Count, "A").End(xlUp).row - 1
    
    ' Replace lines in the array
    Dim i As Integer, ln As Integer
    For i = 1 To lrow
        ln = wks.Range("A1").Offset(i, 1).Value - 1
        repLine(ln) = twospace & wks.Range("A1").Offset(i, 2).Value
    Next
    
    ' Overwrite original JSON
    writeFile.Write Join(repLine, vbNewLine)
    writeFile.Close
    
    ' Clean up
    Set readFile = Nothing
    Set writeFile = Nothing
    Set FSO = Nothing
End Sub

Sub SimJSON()
    ' To keep things simple, this simply replaces certain lines in the JSON instead of creating the JSON from scratch

    ' Declare variables
    Const ForReading = 1    '
    Dim fileToRead As String: fileToRead = Range("SimJSON").Value   ' the path of the file to read
    Dim fileToWrite As String: fileToWrite = fileToRead                 ' the path of a new file (set to be the same as the read file)
    Dim FSO As Object
    Dim readFile As Object      'the file you will READ
    Dim writeFile As Object     'the file you will CREATE (set to be the same as the read file)
    Dim repLine As Variant      'the array of lines you will WRITE
    Dim lrow As Long
    Dim l As Long
    Dim wks As Worksheet
    
    ' Read entire file into an array & close it
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set readFile = FSO.OpenTextFile(fileToRead, ForReading, False)
    repLine = Split(readFile.ReadAll, vbNewLine)
    readFile.Close
    
    Set writeFile = FSO.CreateTextFile(fileToRead, True, False)
    
    Dim dq As String: dq = """"
    Dim escapedDq As String: escapedDq = "\"""
    Dim twospace As String: twospace = "  "
    Dim fourspace As String: fourspace = "    "
    
    ' Set worksheet name and number of rows to loop through
    Set wks = ThisWorkbook.Sheets("GFBattleSimulator")
    lrow = wks.Cells(Rows.Count, "A").End(xlUp).row - 1
    
    ' Replace lines in the array
    Dim i As Integer, ln As Integer
    For i = 1 To lrow
        ln = wks.Range("A1").Offset(i, 1).Value - 1
        repLine(ln) = fourspace & wks.Range("A1").Offset(i, 2).Value
    Next
    
    ' Overwrite original JSON
    writeFile.Write Join(repLine, vbNewLine)
    writeFile.Close
    
    ' Clean up
    Set readFile = Nothing
    Set writeFile = Nothing
    Set FSO = Nothing
End Sub
