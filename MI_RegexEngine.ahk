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
;    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "MatchNum" "SubmatchNum"
;      Alias: "Regex Match", "RegexMatch", "Match", "M" - CASE INSENSITIVE
;      SubmatchNum: 1, 2, 3, ... or *
;      MatchNum: 1, 2, 3, ... or *
;
;    Description:
;      Regex Match reads all data from the column <InputColumn> of table <TableName>,
;      uses RegEx to find the needle specified, and returns the found data to the
;      column <OutputColumn> of table <TableName>.
;
;      If MatchNum = 1 then the 1st match will be returned
;      If MatchNum = 2 then the 2nd match will be returned
;      If MatchNum = * then     all matches will be returned seperated by a |
;
;      If SubmatchNum = 0 then the whole pattern is returned
;      If SubmatchNum = 1 then the 1st captured subpatterns is returned
;      If SubmatchNum = 2 then the 2nd captured subpatterns is returned
;      If SubmatchNum = * then the whole pattern and all captured subpatterns
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
;
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
;  Including?:
;	HWND:{1234567}
;
;MapBasic Window Scripting:
;  If you are making a script in the MapBasic window you can call the Regex Engine with this:
;    UNCOMPILED:
;      run program """C:\Program Files\AutoHotkey\AutoHotkey.exe"" ""C:\Users\jwa\Desktop\MapInfoRegex-master\MI_RegexEngine.ahk"" ""Alias"" ""Params..."""
;      note "Waiting for Regex response..."
;      note "Regex Complete! Execute more code here!"
;
;    COMPILED:
;      run program """<ENGINE-PATH>\MI_RegexEngine.exe"" ""Alias"" ""Params..."""
;      note "Waiting for Regex response..."
;      note "Regex Complete! Execute more code here!"
;
;  The code above will run the RegEx engine, freeze the MapBasic script while the regex engine is running, and unfreeze
;   the running script again after the RegEx engine has completed it's task.
;
;  If you are writing code in MapBasic to be compiled to MBX, use ShellAndWait(). (See: http://stackoverflow.com/a/41250939/6302131 for an implementation of this.)
;
;OUTPUT TO MESSAGE WINDOW:
;  If the output column name is $msg or $message then the change will be outputted to the message window. Useful for experimentation!
;  BUG:
;  If blank lines are found in the column currently there is a bug where these data entries are not printed to the message window.
;  However when updating the MapInfo table everything will work as normal.
;
;
;TODO:
;  Add default GUI.

#Include MI_MultiInstancing.ahk
DEBUGMode := 1

;Get all parameters
p := GetParamArray()

if p.length() = 0 {
	;Open GUI app
	MsgBox, This is currently a command line tool only.
	ExitApp
}

;TRY {
	;***************
	;* SWITCHBOARD *
	;***************
	;Regex Match		;MapInfo_RegexMatch(TableName,InputColumn,OutputColumn,Needle,Options,MatchNum,Count)
	If (p[1] ~= "i)(?:(regex)?\s*match|^m$)") {
		;msgbox, Match
		MapInfo_RegexMatch(p[2],p[3],p[4],p[5],p[6],p[7],p[8])
		GoTo, EndAutotExec
	}
	
	;Regex Replace	;MapInfo_RegexReplace(TableName,InputColumn,OutputColumn,Needle,Options,sReplace)
	if (p[1] ~= "i)(?:(regex)?\s*replace|^r$)") {
		;msgbox, Replace
		MapInfo_RegexReplace(p[2],p[3],p[4],p[5],p[6],p[7])
		GoTo, EndAutotExec
	}
	
	;Regex Format		;MapInfo_RegexFormat(TableName,InputColumn,OutputColumn,Needle,Options,sFormat)
	if (p[1] ~= "i)(?:(regex)?\s*format|^f$)") {
		;msgbox, Format
		MapInfo_RegexFormat(p[2],p[3],p[4],p[5],p[6],p[7])
		GoTo, EndAutotExec
	}
	
	;Return the count of the occurrencees of a given pattern in a string.
	;Regex Count		;MapInfo_RegexCount(TableName,InputColumn,OutputColumn,Needle,Options)
	if (p[1] ~= "i)(?:(regex)?\s*(get\s*)?c(ou)?nt|^c$)") {
		;Msgbox, Count
		MapInfo_RegexCount(p[2],p[3],p[4],p[5],p[6])
		GoTo, EndAutotExec
	}
	
	;Return Pos of found pattern. To find the position of the nth pattern use iPatternNumber = n / ColumnName
	;Regex Pos		;MapInfo_RegexPosition(TableName, InputColumn, OutputColumn, Needle,Options,iPatternNumber=1)
	if (p[1] ~= "i)(?:(regex)?\s*(get\s*)?pos(ition)?|^p$)") {
		;Msgbox, Position
		MapInfo_RegexPos(p[2],p[3],p[4],p[5],p[6],p[7])
		GoTo, EndAutotExec
	}
	
	;Test data extraction
	if (p[1] ~= "i)(?:test|^t$)") {
		MapInfo_Test()
		GoTo, EndAutotExec
	}
	
	;Help - Command Reference
	if (p[1] ~= "i)(?:help|^h$|^\?$)") {
		Msgbox, Todo.
		GoTo, EndAutotExec
	}
	
	Msgbox, % "Unknown command: " . StrJoin(p,",")
	
	
EndAutotExec:
;} CATCH e {
;	Extra := DEBUGMode=1 ? "`n`n" . e : "" 
;	Msgbox,16, MI_RegexEngine.exe - Fatal Error, A fatal error occurred. RegexEngine will quit.%Extra%
;	ExitApp
;}

;Extra functionality for scripters, resume MapBasic Window runtime.
WinClose, MapInfo ahk_class #32770 ahk_exe MapInfoPro.exe, Waiting for Regex response...

ExitApp

;*******************************************************************************
;*                           MAPINFO BASE FUNCTIONS                            *
;*******************************************************************************

MapInfo_RegexMatch(TableName,InputColumn,OutputColumn,Needle,Options,MatchNum,SubMatchNum){
	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)
	
	;Include options:
	Needle := "O" . StrReplace(Options, "O") . ")" . Needle
	
	;Regex Match   ;IMPORTANT: Array_RegexMatch expects object returning needle.
	Array_RegexMatch(data, Needle, MatchNum, SubMatchNum)
	
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
	Array_RegexFormat(data, Needle, sFormat)
	
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
	MI.Do("Print ""Testing " . A_ScriptName . """")

	Loop, % data.Length()
	{
		MI.Do("Print """ data[A_Index] """")
	}
}

;*******************************************************************************
;*                               ARRAY FUNCTIONS                               *
;*******************************************************************************

Array_RegexMatch(byref data, Needle, iMatchNum, iSubmatch){
	;IMPORTANT: Needle is expected to return object!!!
	
	;Loop through all element in array
	Loop, % data.Length()
	{
		dataline := A_Index
		i := 1
		While i {
			oMatch=0
			i := RegexMatch(data[dataline],Needle, oMatch,i)
			i := oMatch.Pos(0) + oMatch.Len(0)
			
			if not i 
				break
			
;			msgbox, % "Index: " dataline ", Data: " data[dataline] ", Needle: " Needle
			
			If (A_Index = iMatchNum) {		;If iMatchNum = int return match(int)
				;If iSubmatch = "*" return all extracted subpatterns
				If (iSubmatch = "*"){
					datapart := Regex_RetAllSubPatterns(oMatch)
				} else {
					datapart := oMatch.Value(iSubmatch)
				}
				
				break ;GoTo, NextArrayElement
			} else if (iMatchNum = "*") {	;If iMatchNum = "*" return all matches
				If (iSubmatch = "*"){
					datapart := datapart = "" ? Regex_RetAllSubPatterns(oMatch) : datapart "|" Regex_RetAllSubPatterns(oMatch)
					;msgbox, % "Datapart: " datapart ", Subs: " Regex_RetAllSubPatterns(oMatch)
				} else {
					datapart := datapart = "" ? oMatch.Value(iSubmatch) :  datapart "|" oMatch.Value(iSubmatch)
				}
			}
;			msgbox, % "Match: " oMatch.value(0) ", iMatchNum: " oMatch.Count() ", Datapart: " datapart
		}
	;NextArrayElement:
		data[A_Index] := datapart
		;msgbox, % datapart
		datapart =
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
		;msgbox, % "Needle: " Needle "`nFormat: " sFormat "`nData: " data[A_Index]
		
		;Get the match from the needle
		i := RegexMatch(data[A_Index],Needle,Match,1)

		;Replace full pattern and all other subpatterns.
		data[A_Index] := RegExReplace(Match,Needle, sFormat)
		
		;msgbox, % "Result: " data[A_Index]
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
		iCount := 1
		LoopIndex := A_Index
		while i<>0 {
			;find potential "Pos"
			i := RegexMatch(data[LoopIndex],Needle, oMatch,i)

			if (iPatternNumber = iCount){
				;before:=data[LoopIndex]
				data[LoopIndex] := i
				;msgbox , % "Before: " before "`nAfter: " data[LoopIndex]
				break
			}
			
			;Setup i for next pattern
			i := oMatch.Pos(0) + oMatch.Len(0)

			;Increment iCount
			iCount++
		}
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
		;OnError Table doesn't exist
		try {
			MI.Do("Fetch Rec " A_Index " from " TableName)
		} catch {
			Msgbox, 48, RegexEngine - Sytax Error, Cannot find table "%TableName%"
			ExitApp
		}
		
		;OnError Column doesn't exist
		try {
			data.push(MI.Eval(TableName "." ColumnName))
		} catch e {
			Msgbox, 48, RegexEngine - Sytax Error, Cannot find column "%ColumnName%"
			ExitApp
		}
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
			MI.Do("Print """ . data[A_Index] . """")
		}

	} else {
		;Get number of rows of table
		iNumRows := MI.Eval("TableInfo(" TableName ",8)")

		;Loop through all rows and set values
		Loop, %iNumRows%
		{
			thisData := data[A_Index]
			MI.Do("Update " TableName " set " ColumnName " = """ thisData """ where rowid = " A_Index)
		}
	}
}

MapInfo_GetApp(){
	If (!MI := GetMapInfo()){
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

StrJoin(obj,delimiter:="",OmitChars:=""){
	string:=obj[1]
	Loop % obj.MaxIndex()-1
		string .= delimiter Trim(obj[A_Index+1],OmitChars)
	return string
}

;Get command line arguments function
;By SKAN https://autohotkey.com/boards/viewtopic.php?t=4357

Args( CmdLine := "", Skip := 0 ) {     ; By SKAN,  http://goo.gl/JfMNpN,  CD:23/Aug/2014 | MD:24/Aug/2014
  Local pArgs := 0, nArgs := 0, A := []
  
  pArgs := DllCall( "Shell32\CommandLineToArgvW", "WStr",CmdLine, "PtrP",nArgs, "Ptr" ) 

  Loop % ( nArgs ) 
     If ( A_Index > Skip ) 
       A[ A_Index - Skip ] := StrGet( NumGet( ( A_Index - 1 ) * A_PtrSize + pArgs ), "UTF-16" )  

Return A,   A[0] := nArgs - Skip,   DllCall( "LocalFree", "Ptr",pArgs )  
}

;Wrapper for SKAN's command arguments function.
GetParamArray(){

	CmdLine := DllCall( "GetCommandLine", "Str" )
	Skip    := ( A_IsCompiled ? 1 : 2 )
	p	:= Args( CmdLine, Skip )
	
	return p
}

GetMITables(){
	Tables := []
	MI:=MapInfo_GetApp()
	Loop, % MI.Eval("numtables()")
	{
		Tables.push(MI.Eval("tableinfo(" . A_Index . ",1)"))
	}
	
	return Tables
}

MsgIndex(arr, index){
_ := arr[index]
msgbox, %index%:%_%
}
