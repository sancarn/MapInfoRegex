;FEATURES:
;Regular Expressions support for MapInfo!
;Support brought for the following functions:
;  Regex Match     - Pattern found; Data to chosen column
;  Regex Replace   - Pattern found; Patterns replaced; Data to chosen column
;  Regex Format    - Pattern and subpatterns found; Reformated with \1 \2 \3 ...; Data to chosen column
;  Regex Count		 - Pattern counted. Count returned to chosen column.
;  Regex Position	 - Pattern found. Position returned to chosen column
;
;SYNTAX OVERVIEW:
;  Regex Match:
;    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "MatchNum" "Count"
;      Alias: "Regex Match", "RegexMatch", "Match", "M" - CASE INSENSITIVE
;      MatchNum: 1, 2, 3, ... or *
;      Count: 1, 2, 3, ... or *
;
;    Description:
;      Regex Match reads all data from the column <InputColumn> of table <TableName>,
;      uses RegEx to find the needle specified, and returns the found data to the
;      column <OutputColumn> of table <TableName>.
;
;      If Count = 1 then the 1st match will be returned
;      If Count = 2 then the 2nd match will be returned
;      If Count = * then     all matches will be returned seperated by a |
;
;      If MatchNum = 0 then the whole pattern is returned
;      If MatchNum = 1 then the 1st captured subpatterns is returned
;      If MatchNum = 2 then the 2nd captured subpatterns is returned
;      If MatchNum = * then the whole pattern and all captured subpatterns
;      are returned seperated by ";"
;
;  Regex Replace:
;    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Replacement"
;      Alias: "Regex Replace", "RegexReplace", "Replace", "M" - CASE INSENSITIVE
;
;
;
;
;
;
;
;
;
;
;

;Get all parameters
p := []
Loop, %0%
{
	_p = %A_Index%
	p.Push(_p)
}

;		;Extra functionality for scripters, pause MapBasic Window runtime.
;		If (A_ScriptName ~= "MBWnd") {
;			;Get MapInfo Application
;			MI := MapInfo_GetApp()
;			MI.Do("Note ""Please wait while RegEx executes..."" ")
;		}

;TODO:
;Params IDEA: TableName,InputColumn,OutputColumn, Needle, NeedleOptions , ....

;Regex Match		;MapInfo_RegexMatch(TableName,InputColumn,OutputColumn,Needle,MatchNum,Count)
If (p[0] ~= "i)(?:(regex)?\s*match|m)") {
	MapInfo_RegexMatch(p[1],p[2],p[3],p[4],p[5],p[6])

;Regex Replace	;MapInfo_RegexReplace(TableName,InputColumn,OutputColumn,Needle,sReplace)
} else if (p[0] ~= "i)(?:(regex)?\s*replace|r)") {
	MapInfo_RegexReplace(p[1],p[2],p[3],p[4],p[5])

;Regex Format		;MapInfo_RegexFormat(TableName,InputColumn,OutputColumn,Needle,sFormat, StartingPositionColumn)
} else if (p[0] ~= "i)(?:(regex)?\s*format|f)") {
	MapInfo_RegexFormat(p[1],p[2],p[3],p[4],p[5])

;Return the count of the occurrencees of a given pattern in a string.
;Regex Count		;MapInfo_RegexCount(TableName,InputColumn,OutputColumn,Needle)
} else if (p[0] ~= "i)(?:(regex)?\s*(get\s*)?c(ou)?nt|c)") {
	MapInfo_RegexCount(p[1],p[2],p[3],p[4])

;Return Pos of found pattern. To find the position of the nth pattern use iPatternNumber = n / ColumnName
;Regex Pos		;MapInfo_RegexPosition(TableName, InputColumn, OutputColumn, Needle, iPatternNumber=1)
} else if (p[0] ~= "i)(?:(regex)?\s*(get\s*)?pos(ition)?|p)") {
	MapInfo_RegexPosition(p[1],p[2],p[3],p[4],p[5])

;Test data extraction
} else if (p[0] ~= "i)(?:test|t)") {
	MapInfo_Test()

;Help - Command Reference
} else if (p[0] ~= "i)(?:help|h|\?)") {
	Msgbox, Help! I need assistance!

} else {
	Msgbox, % "Unknown command: " . StrJoin(p,",")
}

;		;Extra functionality for scripters, resume MapBasic Window runtime.
;		If (A_ScriptName ~= "MBWnd") {
;			WinClose, WinTitle, WinText
;		}
ExitApp

;*******************************************************************************
;*                           MAPINFO BASE FUNCTIONS                            *
;*******************************************************************************

MapInfo_RegexMatch(TableName,InputColumn,OutputColumn,Needle,MatchNum,Count){
	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)

	;Regex Match
	Array_RegexMatch(data, Needle, MatchNum, Count)

	;Set output data
	MapInfo_SetData(TableName, OutputColumn, data)
}

MapInfo_RegexReplace(TableName,InputColumn,OutputColumn,Needle,sReplace){
	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)

	;Regex Match
	Array_RegexReplace(data, Needle, sReplace)

	;Set output data
	MapInfo_SetData(TableName, OutputColumn, data)
}

MapInfo_RegexFormat(TableName,InputColumn,OutputColumn,Needle,sFormat){
	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)

	;Regex Match
	Array_RegexFormat(data, Needle, sFormat)

	;Set output data
	MapInfo_SetData(TableName, OutputColumn, data)
}

MapInfo_RegexCount(TableName,InputColumn,OutputColumn,Needle){
	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)

	;Regex Match
	Array_RegexCount(data, Needle)

	;Set output data
	MapInfo_SetData(TableName, OutputColumn, data)
}

MapInfo_RegexPos(TableName,InputColumn,OutputColumn,Needle,iPatternNumber=1){
	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)

	;Regex Match
	Array_RegexPos(data, Needle, iPatternNumber)

	;Set output data
	MapInfo_SetData(TableName, OutputColumn, data)
}

MapInfo_Test(){
	;Get MapInfo Application
	MI := MapInfo_GetApp()

	TableName := MI.Eval("tableinfo(1,1)")
	InputColumn := "Col1"

	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)

	MI.Do("Print chr$(12)")
	MI.Do("Print ""Testing MapInfoRegex.ahk""")

	Loop, % data.Length()
	{
		MI.Do("Print """ data[A_Index] """")
	}
}

;*******************************************************************************
;*                               ARRAY FUNCTIONS                               *
;*******************************************************************************

Array_RegexMatch(byref data, Needle, MatchNum, Count){
	;Add "O" to Needle
	Needle := Regex_Objectify(Needle)

	;Loop through all element in array
	Loop, % data.Length()
	{
		i := 1
		While i <> 0 {
			i := RegexMatch(data[A_Index],Needle, oMatch,i)
			i := oMatch.Pos(0) + oMatch.Len(0)

			If (A_Index = Count) {
				;If MatchNum = "*" return all extracted subpatterns
				If (MatchNum = "*"){
					datapart := Regex_RetAllSubPatterns(oMatch)
				} else {
					datapart := oMatch.Value(MatchNum)
				}
				GoTo, NextArrayElement
			} else if (Count = "*") {
				If (MatchNum = "*"){
					datapart := datapart "|" Regex_RetAllSubPatterns(oMatch)
				} else {
					datapart := datapart "|" oMatch.Value(MatchNum)
				}
			}
		}
	NextArrayElement:
		data[A_Index] = datapart
	}
}

Array_RegexReplace(byref data, Needle, sReplace){
	;Add "O" to Needle
	Needle := Regex_Objectify(Needle)

	;Loop through all element in array
	Loop, % data.Length()
	{
		data[A_Index] := RegExReplace(data[A_Index],Needle, sReplace)
	}
}

Array_RegexFormat(byref data, Needle, sFormat){
	;Rearranges match/subpatterns into new format. E.G.
	;data := ["abc123","def456","ghi789"]
	;Array_RegexFormat(data,"\w\w\w(\d\d\d)","Suffix: \1")
	;--> data == ["123","456","789"]

	;data := ["abc123","def456","ghi789"]
	;Array_RegexFormat(data,"(\w\w\w)(\d\d\d)","\2\1")
	;--> data == ["123abc","456def","789ghi"]

	;Add "O" to Needle
	Needle := Regex_Objectify(Needle)

	;Loop through all element in array
	Loop, % data.Length()
	{
		;Set ID for later use in next nested loop
		ID := A_Index

		;Get the match from the needle
		i := RegexMatch(data[A_Index],Needle, oMatch,1)

		;Replace full pattern and all other subpatterns.
		Loop, % oMatch.Count() + 1
		{
			;Recall, Loop starts at 1
			index := A_Index -1
			StringReplace, data[ID], data[ID], "\" . index , oMatch.Value(index)
		}
	}
}

Array_RegexCount(byref data, Needle){
	;Add "O" to Needle
	Needle := Regex_Objectify(Needle)

	;Loop through all element in array
	Loop, % data.Length()
	{
		a := RegExReplace(data[A_Index],Needle, "", count)
		data[A_Index] := count
	}
}

Array_RegexPos(byref data, Needle, iPatternNumber){
	;Add "O" to Needle
	Needle := Regex_Objectify(Needle)

	;Loop through all element in array
	Loop, % data.Length()
	{
		i := 1
		count := 1
		while i<>0 {
			;find potential "Pos"
			i := RegexMatch(data[A_Index],Needle, oMatch,i)

			if (iPatternNumber = count){
				data[A_Index] := i
				GoTo NextArrayElement  ;break?
			}

			;Setup i for next pattern
			i := oMatch.Pos(0) + oMatch.Len(0)

			;Increment count
			count := count + 1
		}
NextArrayElement:
	}
}

;*******************************************************************************
;*                              HELPER FUNCTIONS                               *
;*******************************************************************************

MapInfo_GetData(TableName, ColumnName){
	MI := MapInfo_GetApp()
	iNumRows := MI.Eval("TableInfo(" TableName ",8)")

	data := []

	Loop, %iNumRows%
	{
		MI.Do("Fetch Rec " A_Index " from " TableName)
		data.push(MI.Eval(TableName "." ColumnName))
	}

	return data
}

MapInfo_SetData(TableName, ColumnName, data){
	MI := MapInfo_GetApp()
	if (ColumnName ~= "i)\$(msg|message)"){
		;Clear message window
		MI.Do("Print chr$(12)")

		;Loop through data range and print data
		Loop, % data.Length()
		{
			MI.Do("Print """ data[A_Index] """")
		}

	} else {
		;Get number of rows of table
		iNumRows := MI.Eval("TableInfo(" TableName ",8)")

		;Loop through all rows and set values
		Loop, %iNumRows%
		{
			thisData := data[A_Index -1]
			MI.Do("Update " TableName " set " ColumnName " = """ thisData """ where rowid = " A_Index)
		}
	}
}

MapInfo_GetApp(){
	MI := []

	;Try to get MapInfo 64-bit
	MI := ComObjActive("MapInfo.Application.x64")

	;If MI still nothing then get MapInfo 32-Bit
	If (MI = []){
		MI := ComObjActive("MapInfo.Application")
	}

	;If MI still nothing, exit app and error
	If (MI = []){
		Msgbox, MapInfo is not open. MapInfo must be open to use this utility.
		ExitApp
	}
	return MI
}

Regex_RetAllSubPatterns(oMatch){
	;Get first Sub-Pattern
	datapart := oMatch.Value(0)

	;Loop over and combine all Sub-Patterns
	Loop, % oMatch.Count()
	{
		datapart := datapart ";" oMatch.Value(A_Index)
	}
	return datapart
}

Regex_Objectify(Needle){
	RegexMatch(Needle, "((?:i|m|s|x|A|D|J|U|X|P|S|C|O|`n|`r|`a)+?)\)", Match)
	If Not (Match1 ~= ".*O.*") {
		return, "O" . Needle
	} else if (Match1 = "") {
		return, "O)" . Needle
	}
}


;---------------------------
;TestRegexEngine.ahk
;---------------------------
;Command not found: "RegexMatch". RegexMatch = ""
;---------------------------
;OK
;---------------------------
;Command:
;Autohotkey.exe "C:\Users\jwa\Desktop\TestRegexEngine.ahk" "RegexMatch" "A" "A" "B" "i)\w+(\d*)" 1 1
