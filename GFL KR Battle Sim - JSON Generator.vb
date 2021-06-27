' This workbook uses modifications of the code provided on these two sites:
' (1) https://superuser.com/questions/1249898/saving-excel-sheet-as-json-file
' (2) https://stackoverflow.com/questions/19371990/how-do-i-replace-a-string-in-a-line-of-a-text-file-using-filesystemobject-in-vba

Option Explicit

Sub KRSimJSON()
    Call CreateDollJSON
    Call CreateJSON("Fairy", Range("FairyJSONPath").Value)
    Call CreateJSON("Equip", Range("EquipJSONPath").Value)
    Call CreateJSON("HOC", Range("HOCJSONPath").Value)
    Call ChipJSON ' blank file
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
    Dim lrow As Long: lrow = wks.Cells(Rows.Count, "A").End(xlUp).Row
        
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
    Dim lrow As Long: lrow = wks.Cells(Rows.Count, "A").End(xlUp).Row
    
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

Sub ChipJSON()
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim fullFilePath As String: fullFilePath = Range("ChipJSONPath").Value

    Dim fileStream As Object
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 2 'Specify stream type - we want To save text/string data.
    fileStream.Charset = "utf-8" 'Specify charset For the source text data.
    fileStream.Open 'Open the stream And write binary data To the object

    fileStream.WriteText "{" & vbNewLine
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
    lrow = wks.Cells(Rows.Count, "A").End(xlUp).Row - 1
    
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
    lrow = wks.Cells(Rows.Count, "A").End(xlUp).Row - 1
    
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

Sub SaveTeam()
    Application.ScreenUpdating = False
    
    ' Get name for new preset
    NewPresetName = InputBox("Enter a name for the new preset echelon:")
    
    ' Exit if no input or cancelled
    If NewPresetName = vbCancel Or NewPresetName = "" Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    With ThisWorkbook
        Set result = .Sheets("Preset Teams").Range("C:C").Find(NewPresetName)
        
        ' If echelon name doesn't exist, add it
        If result Is Nothing Then
            pasteRow = .Sheets("Preset Teams").Range("D1048576").End(xlUp).Row + 1
            
            .Sheets("Main").Range("EchelonInput").Copy
            .Sheets("Preset Teams").Range("D" & pasteRow).PasteSpecial xlAll
            .Sheets("Preset Teams").Range("C" & pasteRow).Value = NewPresetName
            Application.CutCopyMode = False
            
            ' Adjust list used for dropdown
            pasteRow = .Sheets("Preset Teams").Range("A1048576").End(xlUp).Row + 1
            .Sheets("Preset Teams").Range("A" & pasteRow).Value = NewPresetName
            .Sheets("Preset Teams").Sort.SetRange Range("ListOfPresets")
            .Sheets("Preset Teams").Sort.Apply
            
            ' Set dropdown box value
            .Sheets("Main").Range("Preset").Value = NewPresetName
            
        ' Else overwrite the existing data
        Else
            rowNum = result.Row
            .Sheets("Main").Range("EchelonInput").Copy
            Sheets("Preset Teams").Range("D" & rowNum & ":Y" & rowNum + 4).PasteSpecial xlAll
            Application.CutCopyMode = False
        End If
    End With
    
    Application.ScreenUpdating = True
End Sub

Sub LoadTeam()
    Application.ScreenUpdating = False
    With ThisWorkbook
        presetName = .Sheets("Main").Range("Preset").Value
        Set result = .Sheets("Preset Teams").Range("C:C").Find(presetName)
        
        If result Is Nothing Then
            Application.ScreenUpdating = True
            Exit Sub
        Else
            rowNum = result.Row
            copyRange = .Sheets("Preset Teams").Range("D" & rowNum & ":Y" & rowNum + 4)
            .Sheets("Main").Range("EchelonInput").Value = copyRange
        End If
    End With
    Application.ScreenUpdating = True
End Sub

Sub DeleteTeam()
    Application.ScreenUpdating = False
    ' Ask user to confirm before deletion
    prompt = MsgBox("Pressing 'OK' will delete the preset echelon from the 'Teams' tab. Are you sure you want to continue?", vbYesNo)
    If prompt = vbNo Then Exit Sub

    With ThisWorkbook
        presetName = .Sheets("Main").Range("Preset").Value
        
        ' Delete echelon input
        Set result = .Sheets("Preset Teams").Range("C:C").Find(presetName)
        
        If result Is Nothing Then
            Application.ScreenUpdating = True
            Exit Sub
        Else
            rowNum = result.Row
            .Sheets("Preset Teams").Range("C" & rowNum & ":Y" & rowNum + 4).Delete (xlShiftUp)
            
            ' Delete from dropdown list
            rowNum = .Sheets("Preset Teams").Range("A:A").Find(presetName).Row
            .Sheets("Preset Teams").Range("A" & rowNum).Delete (xlShiftUp)
            
            ' Clear dropdown box
            .Sheets("Main").Range("Preset").Value = ""
        End If
    End With
    Application.ScreenUpdating = False
End Sub
