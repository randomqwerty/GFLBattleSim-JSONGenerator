' JSONCreation Module
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
		Dim lrow As Long: lrow = 8
		
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

' SaveLoadClearInputs Module
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
			Set result = .Sheets("Preset Teams").Range("C:C").Find(NewPresetName, LookAt:=xlWhole)
			
			' If echelon name doesn't exist, add it
			If result Is Nothing Then
				pasteRow = .Sheets("Preset Teams").Range("C1048576").End(xlUp).Row + 5
				.Sheets("Preset Teams").Range("C" & pasteRow).Value = NewPresetName
				
				' Echelon input
				.Sheets("Main").Range("EchelonInput").Copy
				.Sheets("Preset Teams").Range("D" & pasteRow).PasteSpecial xlAll
				
				' Fairy input
				.Sheets("Main").Range("FairyInput").Copy
				.Sheets("Preset Teams").Range("AG" & pasteRow).PasteSpecial xlAll
				
				' Position input
				.Sheets("Main").Range("PositionInput").Copy
				.Sheets("Preset Teams").Range("AM" & pasteRow).PasteSpecial xlAll
				Selection.FormatConditions.Delete
				
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
				Sheets("Preset Teams").Range("D" & rowNum).PasteSpecial xlAll
				
				.Sheets("Main").Range("FairyInput").Copy
				Sheets("Preset Teams").Range("AG" & rowNum).PasteSpecial xlAll
				
				.Sheets("Main").Range("PositionInput").Copy
				Sheets("Preset Teams").Range("AM" & rowNum).PasteSpecial xlAll
				Application.CutCopyMode = False
			End If
		End With
		
		Sheets("Preset Teams").Range("D:AI").FormatConditions.Delete
		Sheets("Preset Teams").Range("D:AI").Validation.Delete
		Application.ScreenUpdating = True
	End Sub

	Sub LoadTeam()
		Application.ScreenUpdating = False
		With ThisWorkbook
			presetName = .Sheets("Main").Range("Preset").Value
			If presetName = "" Then
				Application.ScreenUpdating = True
				Exit Sub
			End If
				
			Set result = .Sheets("Preset Teams").Range("C:C").Find(presetName, LookAt:=xlWhole)
			
			If result Is Nothing Then
				Application.ScreenUpdating = True
				Exit Sub
			Else
				rowNum = result.Row
				copyRange = .Sheets("Preset Teams").Range("D" & rowNum & ":AE" & rowNum + 4)
				.Sheets("Main").Range("EchelonInput").Value = copyRange
				
				copyRange = .Sheets("Preset Teams").Range("AG" & rowNum & ":AK" & rowNum)
				.Sheets("Main").Range("FairyInput").Value = copyRange
				
				copyRange = .Sheets("Preset Teams").Range("AM" & rowNum & ":AO" & rowNum + 2)
				.Sheets("Main").Range("PositionInput").Value = copyRange
			End If
		End With
		Application.ScreenUpdating = True
	End Sub

	Sub DeleteTeam()
		Application.ScreenUpdating = False
		' Ask user to confirm before deletion
		prompt = MsgBox("Pressing 'OK' will delete the preset echelon from the 'Preset Teams' tab. Are you sure you want to continue?", vbYesNo)
		If prompt = vbNo Then Exit Sub

		With ThisWorkbook
			presetName = .Sheets("Main").Range("Preset").Value
			If presetName = "" Then
				Application.ScreenUpdating = True
				Exit Sub
			End If
			
			' Delete echelon input
			Set result = .Sheets("Preset Teams").Range("C:C").Find(presetName, LookAt:=xlWhole)
			
			If result Is Nothing Then
				Application.ScreenUpdating = True
				Exit Sub
			Else
				rowNum = result.Row
				.Sheets("Preset Teams").Range("C" & rowNum & ":AO" & rowNum + 4).Delete (xlShiftUp)
				
				' Delete from dropdown list
				rowNum = .Sheets("Preset Teams").Range("A:A").Find(presetName, LookAt:=xlWhole).Row
				.Sheets("Preset Teams").Range("A" & rowNum).Delete (xlShiftUp)
				
				' Clear dropdown box
				.Sheets("Main").Range("Preset").Value = ""
			End If
		End With
		Application.ScreenUpdating = False
	End Sub

	Sub SaveTeamSF()
		Application.ScreenUpdating = False
		
		' Get name for new preset
		NewPresetName = InputBox("Enter a name for the new preset echelon:", "Preset Saving", ThisWorkbook.Sheets("Main").Range("PresetSF").Value)
		
		' Exit if no input or cancelled
		If NewPresetName = vbCancel Or NewPresetName = "" Then
			Application.ScreenUpdating = True
			Exit Sub
		End If
		
		With ThisWorkbook
			Set result = .Sheets("Preset Teams SF").Range("C:C").Find(NewPresetName, LookAt:=xlWhole)
			
			' If echelon name doesn't exist, add it
			If result Is Nothing Then
				pasteRow = .Sheets("Preset Teams SF").Range("C1048576").End(xlUp).Row + 9
				.Sheets("Preset Teams SF").Range("C" & pasteRow).Value = NewPresetName
				
				' Echelon input
				.Sheets("Main").Range("SFEchelonInput").Copy
				.Sheets("Preset Teams SF").Range("D" & pasteRow).PasteSpecial xlAll
				
				' Custom stats input
				.Sheets("Main").Range("SFCustomStatInput").Copy
				.Sheets("Preset Teams SF").Range("Q" & pasteRow).PasteSpecial xlAll
				
				' Position input
				.Sheets("Main").Range("SFPositionInput").Copy
				.Sheets("Preset Teams SF").Range("S" & pasteRow).PasteSpecial xlAll
				Application.CutCopyMode = False
				
				' Adjust list used for dropdown
				pasteRow = .Sheets("Preset Teams SF").Range("A1048576").End(xlUp).Row + 1
				.Sheets("Preset Teams SF").Range("A" & pasteRow).Value = NewPresetName
				.Sheets("Preset Teams SF").Sort.SetRange Range("ListOfPresetsSF")
				.Sheets("Preset Teams SF").Sort.Apply
				
				' Set dropdown box value
				.Sheets("Main").Range("PresetSF").Value = NewPresetName
				
			' Else overwrite the existing data
			Else
				rowNum = result.Row
				.Sheets("Main").Range("SFEchelonInput").Copy
				Sheets("Preset Teams SF").Range("D" & rowNum).PasteSpecial xlAll
				
				.Sheets("Main").Range("SFCustomStatInput").Copy
				Sheets("Preset Teams SF").Range("Q" & rowNum).PasteSpecial xlAll
				
				.Sheets("Main").Range("SFPositionInput").Copy
				Sheets("Preset Teams SF").Range("S" & rowNum).PasteSpecial xlAll
				Application.CutCopyMode = False
			End If
		End With
		
		Sheets("Preset Teams SF").Range("D:S").FormatConditions.Delete
		Sheets("Preset Teams SF").Range("D:S").Validation.Delete
		Application.ScreenUpdating = True
	End Sub

	Sub LoadTeamSF()
		Application.ScreenUpdating = False
		With ThisWorkbook
			presetName = .Sheets("Main").Range("PresetSF").Value
			If presetName = "" Then
				Application.ScreenUpdating = True
				Exit Sub
			End If
			
			Set result = .Sheets("Preset Teams SF").Range("C:C").Find(presetName, LookAt:=xlWhole)
			
			If result Is Nothing Then
				Application.ScreenUpdating = True
				Exit Sub
			Else
				rowNum = result.Row
				copyRange = .Sheets("Preset Teams SF").Range("D" & rowNum & ":O" & rowNum + 8)
				.Sheets("Main").Range("SFEchelonInput").Value = copyRange
				
				copyRange = .Sheets("Preset Teams SF").Range("Q" & rowNum & ":Q" & rowNum + 8)
				.Sheets("Main").Range("SFCustomStatInput").Value = copyRange
				
				copyRange = .Sheets("Preset Teams SF").Range("S" & rowNum & ":U" & rowNum + 2)
				.Sheets("Main").Range("SFPositionInput").Value = copyRange
			End If
		End With
		Application.ScreenUpdating = True
	End Sub

	Sub DeleteTeamSF()
		Application.ScreenUpdating = False
		' Ask user to confirm before deletion
		prompt = MsgBox("Pressing 'OK' will delete the preset echelon from the 'Preset Teams SF' tab. Are you sure you want to continue?", vbYesNo)
		If prompt = vbNo Then Exit Sub

		With ThisWorkbook
			presetName = .Sheets("Main").Range("PresetSF").Value
			If presetName = "" Then
				Application.ScreenUpdating = True
				Exit Sub
			End If
			
			' Delete echelon input
			Set result = .Sheets("Preset Teams SF").Range("C:C").Find(presetName, LookAt:=xlWhole)
			
			If result Is Nothing Then
				Application.ScreenUpdating = True
				Exit Sub
			Else
				rowNum = result.Row
				.Sheets("Preset Teams SF").Range("C" & rowNum & ":U" & rowNum + 8).Delete (xlShiftUp)
				
				' Delete from dropdown list
				rowNum = .Sheets("Preset Teams SF").Range("A:A").Find(presetName, LookAt:=xlWhole).Row
				.Sheets("Preset Teams SF").Range("A" & rowNum).Delete (xlShiftUp)
				
				' Clear dropdown box
				.Sheets("Main").Range("PresetSF").Value = ""
			End If
		End With
		Application.ScreenUpdating = False
	End Sub

' DataUpdater Module
	Sub UpdateData()
		With ThisWorkbook
			Call Updater(.Sheets("Misc Inputs").Range("gunScript").Value, "gun", "'Doll Data'!A:A", .Sheets("Doll Data"))
			Call Updater(.Sheets("Misc Inputs").Range("equipScript").Value, "equip", "'Equip Data'!A:A", .Sheets("Equip Data"))
			Call Updater(.Sheets("Misc Inputs").Range("sfScript").Value, "sangvis", "'SF Data'!A:A", .Sheets("SF Data"))
			Call Updater(.Sheets("Misc Inputs").Range("sfResolScript").Value, "sangvis_resolution", "'SF Resolution'!A:A", .Sheets("SF Resolution"))
		End With
	End Sub

	Sub Updater(mScript As String, qryName As String, compareRange As String, outputSheet As Worksheet)
	' mScript: M script used in the actual PowerQuery
	' qryName: name of new query that will contain the M script
	' compareRange: range to compare new data against (only new data will be copied)
	' outputSheet: sheet to paste new data to (will be appended at the end)

		Dim qry As WorkbookQuery
		Dim tempSheet As Worksheet
		
		' Add query and temporary sheet for data
		Set qry = ThisWorkbook.Queries.Add(qryName, mScript)
		Set tempSheet = ThisWorkbook.Sheets.Add
		
		' Load data from query
		With tempSheet.ListObjects.Add( _
			SourceType:=0, _
			Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qry.Name, _
			Destination:=Range("$A$1")).QueryTable
			
				.CommandType = xlCmdDefault
				.CommandText = Array("SELECT * FROM [" & qry.Name & "]")
				.RowNumbers = False
				.FillAdjacentFormulas = False
				.PreserveFormatting = True
				.RefreshOnFileOpen = False
				.BackgroundQuery = True
				.RefreshStyle = xlInsertDeleteCells
				.SavePassword = False
				.SaveData = True
				.AdjustColumnWidth = True
				.RefreshPeriod = 0
				.PreserveColumnInfo = False
				.Refresh BackgroundQuery:=False
		End With

		' Filter data for new entries only
		newColumn = tempSheet.Range("A1").End(xlToRight).Column + 1
		tempSheet.Cells(1, newColumn).Value = "Check"
		tempSheet.Cells(2, newColumn).Formula = "=COUNTIF(" & compareRange & ", A2)"
		tempSheet.ListObjects(1).Range.AutoFilter Field:=newColumn, Criteria1:=0
		
		' Check if there is data to copy after filtering
		On Error Resume Next
		Set rngFiltered = Nothing
		Set rngFiltered = tempSheet.ListObjects(1).DataBodyRange.SpecialCells(xlCellTypeVisible)
		On Error GoTo 0
		
		' Copy new data
		If Not (rngFiltered Is Nothing) Then
			endRow = outputSheet.Range("A1048576").End(xlUp).Row
			tempSheet.ListObjects(1).DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
			outputSheet.Range("A" & endRow + 1).PasteSpecial xlPasteValues
		End If
		
		' Delete temp sheet, query, and workbook connection
		Application.DisplayAlerts = False
		tempSheet.Delete
		qry.Delete
		ThisWorkbook.Connections(1).Delete
		Application.DisplayAlerts = False
	End Sub

' WorkbookUpdater Module
	' Code references:
	'   https://wellsr.com/vba/2018/excel/download-files-with-vba-urldownloadtofile/
	'   https://www.extendoffice.com/documents/excel/3236-excel-delete-current-file-workbook.html

	Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
		Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
		ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

	Sub UpdateWorkbook()
		If MsgBox("This workbook will be deleted and replaced with the latest version from GitHub. Continue?", vbYesNo) = vbNo Then Exit Sub
		transferPresets = MsgBox("Would you like to transfer echelon presets to the new workbook?", vbYesNo)
		downloadUserinfo = MsgBox("Would you like to download the latest userinfo.json?", vbYesNo)
		
		' Store this workbook's name in a variable so it can be used later
		oldName = ThisWorkbook.FullName
		
		' Temporarily rename this workbook and delete the old name so it can be replaced
		ThisWorkbook.SaveAs Filename:=Replace(oldName, ".xlsm", "_temp.xlsm")
		Kill oldName
		
		' Download new version from GitHub
		newFileURL = "https://github.com/randomqwerty/GFLBattleSim-JSONGenerator/raw/main/GFL%20KR%20Battle%20Sim%20-%20JSON%20Generator.xlsm"
		URLDownloadToFile 0, newFileURL, oldName, 0, 0
		
		' Download userinfo.json from GitHub if user said yes
		If downloadUserinfo = vbYes Then
			newFileURL = "https://raw.githubusercontent.com/randomqwerty/GFLBattleSim-JSONGenerator/main/userinfo.json"
			URLDownloadToFile 0, newFileURL, ThisWorkbook.Path & "\Preset\userinfo.json", 0, 0
		End If
	
		' Open new workbook and delete this workbook
		Set newBook = Workbooks.Open(oldName)
		ThisWorkbook.Saved = True
		ThisWorkbook.ChangeFileAccess xlReadOnly
		Kill ThisWorkbook.FullName
		
		' Transfer echelon presets
		If transferPresets = vbYes Then
			Application.Calculation = xlCalculationManual
			' G&K
			newBook.Sheets("Preset Teams").Range("A4:AO1000").Clear
			ThisWorkbook.Sheets("Preset Teams").Range("A4:AO1000").Copy
			newBook.Sheets("Preset Teams").Range("A4").PasteSpecial
			
			' SF
			newBook.Sheets("Preset Teams SF").Range("A4:AO1000").Clear
			ThisWorkbook.Sheets("Preset Teams SF").Range("A4:AO1000").Copy
			newBook.Sheets("Preset Teams SF").Range("A4").PasteSpecial
			Application.CutCopyMode = False
			Application.Calculation = xlCalculationAutomatic
		End If
		
		' Load XML file from GitHub
		Set XDoc = CreateObject("MSXML2.DOMDocument")
		XDoc.async = False: XDoc.validateOnParse = False
		XDoc.Load ("https://raw.githubusercontent.com/randomqwerty/GFLBattleSim-JSONGenerator/main/change.xml")
		
		' Get data and display message
		Set updateDate = XDoc.getElementsByTagName("UpdateDate")
		Set updateMessage = XDoc.getElementsByTagName("Message")
		MsgBox ("Successfully updated workbook to version as of " & updateDate(0).Text & ":" & vbNewLine & vbNewLine & updateMessage(0).Text & vbNewLine & vbNewLine & "For a more complete list of changes, please see the commit history on the repo.")
		
		' Close this workbook now that it has been replaced
		ThisWorkbook.Close SaveChanges:=False
	End Sub