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
;    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "MatchNum" "Count"
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
;    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "Replacement"
;      Alias: "Regex Replace", "RegexReplace", "Replace", "R" - CASE INSENSITIVE
;      Replacement - The string to be substituted for each match, which is plain text (not a regular
;       expression). It may include backreferences like $1, which brings in the substring from Haystack
;  		  that matched the first subpattern. The simplest backreferences are $0 through $9, where $0 is
;  		  the substring that matched the entire pattern, $1 is the substring that matched the first subpattern,
;  		  $2 is the second, and so on. For backreferences above 9 (and optionally those below 9), enclose the
;  		  number in braces; e.g. ${10}, ${11}, and so on. For named subpatterns, enclose the name in braces;
;  		  e.g. ${SubpatternName}. To specify a literal $, use $$ (this is the only character that needs such
;  		  special treatment; backslashes are never needed to escape anything).
;
;       To convert the case of a subpattern, follow the $ with one of the following characters: U or u (uppercase),
;       L or l (lowercase), T or t (title case, in which the first letter of each word is capitalized but all others are
;       made lowercase). For example, both $U1 and $U{1} transcribe an uppercase version of the first subpattern.

;       Nonexistent backreferences and those that did not match anything in Haystack -- such as one of the
;       subpatterns in "(abc)|(xyz)" -- are transcribed as empty strings.
;
;  Regex Format:
;    Description:
;      Regex format takes the first match found given by needles and reformats it with the pattern specified
;				by sFormat. sFormat follows the same rules as "Replacement" of RegexReplace above.
;    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "sFormat"
;      Alias: "Regex Format", "RegexFormat", "Format", "F" - CASE INSENSITIVE
;
;  Regex Count:
;    Description:
;      Regex Count counts and returns the number of occurrences of a pattern.
;    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options"
;      Alias: "Regex Count", "Count", "Cnt", "C" (OTHERS: "(?:(regex)?\s*(get\s*)?c(ou)?nt|c)")
;
;  Regex Position:
;    Description:
;      Regex Position returns the Nth position of a pattern.
;    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "N"
;      Alias: "Regex Position", "Position", "Pos", "P" (OTHERS: "(?:(regex)?\s*(get\s*)?pos(ition)?|p)")
;      N: Default = 1
;
;Regex Options:
;
;  See table of: https://autohotkey.com/docs/misc/RegEx-QuickRef.htm#Options
;
;TODO:
;  Add default GUI.

;Get all parameters
p := []
Loop, %0%
{
	_p = %A_Index%
	p.Push(_p)
}

if p.length() = 0 {
	;Open GUI app
	MsgBox, This is currently a command line tool only.
}

;		;Extra functionality for scripters, pause MapBasic Window runtime.
;		If (A_ScriptName ~= "MBWnd") {
;			;Get MapInfo Application
;			MI := MapInfo_GetApp()
;			MI.Do("Note ""Please wait while RegEx executes..."" ")
;		}

;TODO:
;Params IDEA: TableName,InputColumn,OutputColumn, Needle, NeedleOptions , ....

;Regex Match		;MapInfo_RegexMatch(TableName,InputColumn,OutputColumn,Needle,Options,MatchNum,Count)
If (p[0] ~= "i)(?:(regex)?\s*match|m)") {
	MapInfo_RegexMatch(p[1],p[2],p[3],p[4],p[5],p[6],p[7])

;Regex Replace	;MapInfo_RegexReplace(TableName,InputColumn,OutputColumn,Needle,Options,sReplace)
} else if (p[0] ~= "i)(?:(regex)?\s*replace|r)") {
	MapInfo_RegexReplace(p[1],p[2],p[3],p[4],p[5],p[6])

;Regex Format		;MapInfo_RegexFormat(TableName,InputColumn,OutputColumn,Needle,Options,sFormat)
} else if (p[0] ~= "i)(?:(regex)?\s*format|f)") {
	MapInfo_RegexFormat(p[1],p[2],p[3],p[4],p[5],p[6])

;Return the count of the occurrencees of a given pattern in a string.
;Regex Count		;MapInfo_RegexCount(TableName,InputColumn,OutputColumn,Needle,Options)
} else if (p[0] ~= "i)(?:(regex)?\s*(get\s*)?c(ou)?nt|c)") {
	MapInfo_RegexCount(p[1],p[2],p[3],p[4],p[5])

;Return Pos of found pattern. To find the position of the nth pattern use iPatternNumber = n / ColumnName
;Regex Pos		;MapInfo_RegexPosition(TableName, InputColumn, OutputColumn, Needle,Options,iPatternNumber=1)
} else if (p[0] ~= "i)(?:(regex)?\s*(get\s*)?pos(ition)?|p)") {
	MapInfo_RegexPos(p[1],p[2],p[3],p[4],p[5],p[6])

;Test data extraction
} else if (p[0] ~= "i)(?:test|t)") {
	MapInfo_Test()

;Help - Command Reference
} else if (p[0] ~= "i)(?:help|h|\?)") {
	Msgbox, Todo.

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

MapInfo_RegexMatch(TableName,InputColumn,OutputColumn,Needle,Options,MatchNum,Count){
	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)

	;Include options:
	Needle := "O" . StrReplace(Options, "O") . ")" . Needle

	;Regex Match   ;IMPORTANT: Array_RegexMatch expects object returning needle.
	Array_RegexMatch(data, Needle, MatchNum, Count)

	;Set output data
	MapInfo_SetData(TableName, OutputColumn, data)
}

MapInfo_RegexReplace(TableName,InputColumn,OutputColumn,Needle,Options,sReplace){
	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)

	;Include options:
	Needle := "O" . StrReplace(Options, "O") . ")" . Needle

	;Regex Match   ;IMPORTANT: Array_RegexReplace expects object returning needle.
	Array_RegexReplace(data, Needle, sReplace)

	;Set output data
	MapInfo_SetData(TableName, OutputColumn, data)
}

MapInfo_RegexFormat(TableName,InputColumn,OutputColumn,Needle,Options,sFormat){
	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)

	;Include options:
	Needle := StrReplace(Options, "O") . ")" . Needle

	;Regex Match   ;IMPORTANT: Array_RegexCount expects NON-object returning needle.
	Array_RegexFormat(data, Needle, Options, sFormat)

	;Set output data
	MapInfo_SetData(TableName, OutputColumn, data)
}

MapInfo_RegexCount(TableName,InputColumn,OutputColumn,Needle,Options){
	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)

	;Include options:
	Needle := "O" . StrReplace(Options, "O") . ")" . Needle

	;Regex Match   ;IMPORTANT: Array_RegexCount expects object returning needle.
	Array_RegexCount(data, Needle)

	;Set output data
	MapInfo_SetData(TableName, OutputColumn, data)
}

MapInfo_RegexPos(TableName,InputColumn,OutputColumn,Needle, Options,iPatternNumber=1){
	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)

	;Include options:
	Needle := "O" . StrReplace(Options, "O") . ")" . Needle

	;Regex Match   ;IMPORTANT: Array_RegexPos expects object returning needle.
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
	;IMPORTANT: Needle is expected to return object!!!

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
	;IMPORTANT: Needle is expected to return object!!!

	;Loop through all element in array
	Loop, % data.Length()
	{
		data[A_Index] := RegExReplace(data[A_Index],Needle, sReplace)
	}
}

Array_RegexFormat(byref data, Needle, sFormat){
	;IMPORTANT Needle expected to be non-object

	;Rearranges match/subpatterns into new format. E.G.
	;data := ["abc123","def456","ghi789"]
	;Array_RegexFormat(data,"\w\w\w(\d\d\d)","Suffix: \1")
	;--> data == ["123","456","789"]

	;data := ["abc123","def456","ghi789"]
	;Array_RegexFormat(data,"(\w\w\w)(\d\d\d)","\2\1")
	;--> data == ["123abc","456def","789ghi"]

	;OBJECTIFY NOT NEEDED
	;Add "O" to Needle
	;Needle := Regex_Objectify(Needle)

	;Loop through all element in array
	Loop, % data.Length()
	{
		;Get the match from the needle
		i := RegexMatch(data[A_Index],Needle,Match,1)

		;Replace full pattern and all other subpatterns.
		data[A_Index] = RegExReplace(Match,Needle, sFormat)
	}
}

Array_RegexCount(byref data, Needle){
	;IMPORTANT: Needle is expected to return object!!!

	;Loop through all element in array
	Loop, % data.Length()
	{
		a := RegExReplace(data[A_Index],Needle, "", count)
		data[A_Index] := count
	}
}

Array_RegexPos(byref data, Needle, iPatternNumber){
	;IMPORTANT: Needle is expected to return object!!!

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

;DEPRECATED?
;Regex_Objectify(Needle){
;	RegexMatch(Needle, "((?:i|m|s|x|A|D|J|U|X|P|S|C|O|`n|`r|`a)+?)\)", Match)
;	If Not (Match1 ~= ".*O.*") {
;		return, "O" . Needle
;	} else if (Match1 = "") {
;		return, "O)" . Needle
;	}
;}
