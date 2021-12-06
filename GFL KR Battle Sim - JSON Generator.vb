' This workbook uses modifications of the code provided on these two sites:
' (1) https://superuser.com/questions/1249898/saving-excel-sheet-as-json-file
' (2) https://stackoverflow.com/questions/19371990/how-do-i-replace-a-string-in-a-line-of-a-text-file-using-filesystemobject-in-vba

Option Explicit

Sub KRSimJSON()
    Call CreateJSON_A("Fairy", Range("FairyJSONPath").Value)
    Call CreateJSON_A("Equip", Range("EquipJSONPath").Value)
    Call CreateJSON_A("HOC", Range("HOCJSONPath").Value)
    
    Call CreateJSON_B("Doll", Range("DollJSONPath").Value)
    Call CreateJSON_B("SF", Range("SFJSONPath").Value)
    Call CreateJSON_B("SF Chip", Range("SFChipJSONPath").Value)
    
    Call CreateJSON_C("Mission_Act_Info", Range("MissionJSON").Value, 2)
    Call CreateJSON_C("GFBattleSimulator", Range("SimJSON").Value, 4)
    Call CreateJSON_C("Spot_Act_Info_Day", Range("SpotDayJSON").Value, 8)
    Call CreateJSON_C("Spot_Act_Info_Night", Range("SpotNightJSON").Value, 8)
    
    Call CreateSFTeamJSON("SF Team", Range("SFTeamJSONPath").Value) ' ugly mess
    
    Call ChipJSON ' blank file
    MsgBox "Process completed."
End Sub


Public Sub CreateJSON_A(sheetName As String, fullFilePath As String)
' JSONs that start like:
'     [
'        "1": {

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


Public Sub CreateJSON_B(sheetName As String, fullFilePath As String)
' JSONs that start like:
'     [
'        {
'           "id":

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

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
    fileStream.WriteText "["
    
    Dim j As Integer
    Dim cellvalue As String
    
    ' Loop through rows and columns
    For j = 2 To lrow
        ' Only executes if the row has a non-blank ID
        If wks.Cells(j, 1).Value <> "" Then
            
            For i = 1 To lcolumn
                If i = 1 Then
                    fileStream.WriteText vbNewLine & twospace & "{" & vbNewLine
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
        End If
    Next j
    
    fileStream.WriteText vbNewLine & "]"
    fileStream.SaveToFile fullFilePath, 2 'Save binary data To disk
End Sub

Sub CreateJSON_C(sheetName As String, fileToRead As String, numSpaces As Integer)
' Replaces lines in a given JSON, rather than generate a JSON from scratch

    ' Declare variables
    Const ForReading = 1    '
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
    Dim eightspace As String: eightspace = "        "
    Dim spacesUsed As String
    
    If numSpaces = 2 Then
        spacesUsed = twospace
    ElseIf numSpaces = 4 Then
        spacesUsed = fourspace
    ElseIf numSpaces = 8 Then
        spacesUsed = eightspace
    End If
    
    ' Set worksheet name and number of rows to loop through
    Set wks = ThisWorkbook.Sheets(sheetName)
    lrow = wks.Cells(Rows.Count, "A").End(xlUp).Row - 1
    
    ' Replace lines in the array
    Dim i As Integer, ln As Integer
    For i = 1 To lrow
        ln = wks.Range("A1").Offset(i, 1).Value - 1
        repLine(ln) = spacesUsed & wks.Range("A1").Offset(i, 2).Value
    Next
    
    ' Overwrite original JSON
    writeFile.Write Join(repLine, vbNewLine)
    writeFile.Close
    
    ' Clean up
    Set readFile = Nothing
    Set writeFile = Nothing
    Set FSO = Nothing
End Sub

Public Sub CreateSFTeamJSON(sheetName As String, fullFilePath As String)
' Formatted completely different the rest...

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    Dim fileStream As Object
    Set fileStream = CreateObject("ADODB.Stream")
    fileStream.Type = 2 'Specify stream type - we want To save text/string data.
    fileStream.Charset = "utf-8" 'Specify charset For the source text data.
    fileStream.Open 'Open the stream And write binary data To the object

    Dim wkb As Workbook: Set wkb = ThisWorkbook
    Dim wks As Worksheet: Set wks = wkb.Sheets(sheetName)

    Dim lcolumn As Long: lcolumn = 7
    Dim lrow As Long: lrow = 7
    
    Dim titles() As String
    ReDim titles(lcolumn)
    
    Dim dq As String: dq = """"
    Dim escapedDq As String: escapedDq = "\"""
    Dim twospace As String: twospace = "  "
    Dim fourspace As String: fourspace = "    "
    Dim eightspace As String: eightspace = "        "
    
    ' Define array of column titles
    Dim i As Integer
    For i = 1 To lcolumn
        titles(i) = wks.Cells(1, i)
    Next i
    
    fileStream.WriteText "["
    
    ' Loop through rows and columns
    Dim j As Integer
    Dim cellvalue As String
    For j = 2 To lrow
        If wks.Cells(j, 1).Value <> "" Then
            fileStream.WriteText vbNewLine & twospace & "{" & vbNewLine
            For i = 1 To lcolumn
                
                cellvalue = Replace(wks.Cells(j, i), dq, escapedDq)
                
                If titles(i) <> "info" Then
                    fileStream.WriteText fourspace & dq & titles(i) & dq & ": " & dq & cellvalue & dq
                Else
                    fileStream.WriteText fourspace & dq & titles(i) & dq & ": {"
                End If
                
                If i <> lcolumn Then
                    fileStream.WriteText ","
                End If
                
                fileStream.WriteText vbNewLine
                
                If titles(i) = "info" Then
                    Dim k As Integer
                    If j = 2 Then
                        For k = 1 To 9
                            fileStream.WriteText fourspace & twospace & dq & k & dq & ": {" & vbNewLine
                            fileStream.WriteText eightspace & dq & "sangvis_with_user_id" & dq & ": " & wks.Cells(k + 1, "J") & "," & vbNewLine
                            fileStream.WriteText eightspace & dq & "position" & dq & ": " & wks.Cells(k + 1, "K") & vbNewLine
                            fileStream.WriteText fourspace & twospace & "}"
                            
                            If k <> 9 Then fileStream.WriteText ","
                            fileStream.WriteText vbNewLine
                        Next k
                    Else
                        For k = 1 To 9
                            fileStream.WriteText fourspace & twospace & dq & k & dq & ": {" & vbNewLine
                            
                            If k = 1 Then
                                fileStream.WriteText eightspace & dq & "sangvis_with_user_id" & dq & ": " & wks.Cells(j + 8, "J") & "," & vbNewLine
                                fileStream.WriteText eightspace & dq & "position" & dq & ": " & wks.Cells(j + 8, "K") & vbNewLine
                            Else
                                fileStream.WriteText eightspace & dq & "sangvis_with_user_id" & dq & ": 0," & vbNewLine
                                fileStream.WriteText eightspace & dq & "position" & dq & ": 0" & vbNewLine
                            End If
                            
                            fileStream.WriteText fourspace & twospace & "}"
                            If k <> 9 Then fileStream.WriteText ","
                            fileStream.WriteText vbNewLine
                        Next k
                    End If
                    fileStream.WriteText fourspace & "}" & vbNewLine
                End If
            Next i
            
            fileStream.WriteText twospace & "}"
            
            If j <> lrow Then
                fileStream.WriteText ","
            End If
        End If
    Next j
    fileStream.WriteText vbNewLine & "]"
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

Sub ClearInput()
    With ThisWorkbook.Sheets("Main")
        .Range("EchelonInput").ClearContents
        .Range("CustomStatInput").ClearContents
        .Range("FairyInput").ClearContents
        .Range("PositionInput").ClearContents
        
        .Range("SFEchelonInput").ClearContents
        .Range("SFCustomStatInput").ClearContents
        .Range("SFPositionInput").ClearContents
        
        .Range("HOCSelection").ClearContents
        .Range("SFHOCSelection").ClearContents
        .Range("StrategyInput").ClearContents
        .Range("DebuffSelection").ClearContents
    End With
End Sub
Sub SaveTeam()
    Application.ScreenUpdating = False
    
    ' Get name for new preset
    NewPresetName = InputBox("Enter a name for the new preset echelon:", "Preset Saving", ThisWorkbook.Sheets("Main").Range("Preset").Value)
    
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
            .Sheets("Preset Teams").Range("C" & pasteRow).Value = NewPresetName
            
            ' Echelon input
            .Sheets("Main").Range("EchelonInput").Copy
            .Sheets("Preset Teams").Range("D" & pasteRow).PasteSpecial xlAll
            
            ' Fairy input
            .Sheets("Main").Range("FairyInput").Copy
            .Sheets("Preset Teams").Range("AA" & pasteRow).PasteSpecial xlAll
            
            ' Position input
            .Sheets("Main").Range("PositionInput").Copy
            .Sheets("Preset Teams").Range("AG" & pasteRow).PasteSpecial xlAll
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
            
            .Sheets("Main").Range("FairyInput").Copy
            Sheets("Preset Teams").Range("AA" & rowNum & ":AE" & rowNum).PasteSpecial xlAll
            
            .Sheets("Main").Range("PositionInput").Copy
            Sheets("Preset Teams").Range("AG" & rowNum & ":AY" & rowNum + 2).PasteSpecial xlAll
            
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
            
            copyRange = .Sheets("Preset Teams").Range("AA" & rowNum & ":AE" & rowNum)
            .Sheets("Main").Range("FairyInput").Value = copyRange
            
            copyRange = .Sheets("Preset Teams").Range("AG" & rowNum & ":AI" & rowNum + 2)
            .Sheets("Main").Range("PositionInput").Value = copyRange
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
            .Sheets("Preset Teams").Range("C" & rowNum & ":AI" & rowNum + 4).Delete (xlShiftUp)
            
            ' Delete from dropdown list
            rowNum = .Sheets("Preset Teams").Range("A:A").Find(presetName).Row
            .Sheets("Preset Teams").Range("A" & rowNum).Delete (xlShiftUp)
            
            ' Clear dropdown box
            .Sheets("Main").Range("Preset").Value = ""
        End If
    End With
    Application.ScreenUpdating = False
End Sub