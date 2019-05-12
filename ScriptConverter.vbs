'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2009
'
' NAME: FileScriptConverter v0.1
' 
' AUTHOR: Andrew Meyercord 
' DATE  : 4/4/2010
'
' Copyright (C) 2019 Andrew Meyercord
'
'    This program is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program.  If not, see <https://www.gnu.org/licenses/>.
'==========================================================================

Option Explicit

	Dim objExcel, fso, tsInput, tsOutput, line, char, tcode, os, dialog, nextline, location, synopsis, row
	Dim ColorLoc, ColorDialog, objExcelWB, objExcelWS

	'ColorLoc = "&HB7DEE8" 'Location color
	ColorLoc = "&HE8DEB7" 'Location color
	ColorDialog = "&HFFFFFF"
	Set fso = CreateObject("Scripting.FileSystemObject")

	Set tsInput = fso.opentextfile(fso.GetParentFolderName(wscript.ScriptFullName)+"\script.txt",1)

	'Set tsOutput = fso.CreateTextFile("c:\temp\converted.txt", True)
	Set objExcel = CreateObject("Excel.Application")
	objExcel.Visible = True
	set objExcelWB = objExcel.Workbooks.Add
	set objExcelWS = objExcelWB.Worksheets(1)
	row = 1

	While Not tsInput.AtEndOfStream
		location = ""
		synopsis = ""
		os = "" 'Chr(9)
		nextline = False
		char = ""
		tcode = "" 'Chr(9)
		dialog = ""
		line = ""
		Do until len(line)>0
			line = tsInput.Readline
		Loop
		If Left(ucase(line),3) = "INT" Or Left(ucase(line),3) = "EXT" then
			location = line
			line = ""
			'tsOutput.Writeline Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + location
			objExcelWS.cells(row, 6).value = location
			objExcelWS.Rows(row).Interior.Color = CLng(ColorLoc)
			row = row + 1
			do until len(line)>0
				line = tsInput.Readline
			Loop
		ElseIf Left(ucase(line),4) = "O.S." Then 
			os = "X" '+ os
			line = Right(line, Len(line)-5)
		End if
		Select Case ucase(line)
		Case "ASH"
			char = "A"
		Case "BRETT"
			char ="B"
		Case "DALLAS"
			char = "D"
		Case "KANE"
			char = "K"
		Case "LAMBERT"
			char = "L"
		Case "PARKER"
			char = "P"
		Case "RIPLEY"
			char ="R"
		Case "MOTHER's VOICE"
			char = "M"
		Case Else
			synopsis = line
			line =""
'			tsOutput.Writeline Chr(9) + Chr(9) + Chr(9) + Chr(9) + Chr(9) + synopsis
			objExcelWS.cells(row, 6).value = synopsis
			objExcelWS.Rows(row).Interior.Color = CLng(ColorLoc)
			row = row + 1
			nextline = True
		End Select
		If Not nextline Then
			'char = char + Chr(9)
			line = ""
			do until len(line)>0
				line = tsInput.Readline
			Loop
			If (left(line,1) = "(") And (Right(line,1) = ")") Then
				If isnumeric(Mid(line,2,1)) Then
					tcode = mid(line,2,Len(line)-2)' + Chr(9)
				Else
					dialog = line + " "
				End if
				line = ""
				do until len(line)>0
					line = tsInput.Readline
				Loop
			End If
			dialog = dialog + line
			objExcelWS.cells(row, 2).value = tcode
			objExcelWS.cells(row, 4).value = os
			objExcelWS.cells(row, 5).value = char
			objExcelWS.cells(row, 6).value = dialog
			objExcelWS.Rows(row).Interior.Color = CLng(ColorDialog)
			row = row + 1
'			tsOutput.Writeline Chr(9) + tcode + Chr(9) + os + char + dialog
		End If
	Wend
	tsInput.Close
	'tsOutput.Close
	with objExcelWS
		.range(.Cells(1, 1), .Cells(.rows.count, .columns.count)).borders.linestyle=1
	End with