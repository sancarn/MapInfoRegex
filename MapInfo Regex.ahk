Operation = %1%

;Extra functionality for scripters, pause MapBasic Window runtime.
If (A_ScriptName ~= "MBWnd") {
	;Get MapInfo Application
	MI := MapInfo_GetApp()
	MI.Do("Note ""Please wait while RegEx executes..."" ")
}

If (Operation ~= "i)Regex\s*Match") {
	;
	;TableName	:= %2%
	;InputColumn	:= %3%
	;OutputColumn	:= %4%
	;Needle		:= %5%
	;MatchNum	:= %6%
	;Count		:= %7%
	MapInfo_RegexMatch(%2%,%3%,%4%,%5%,%6%,%7%)

} else if (Operation ~= "i)Test\s*Extract") {
	MapInfo_Test()

} else {
	Msgbox, Command not found: "%1%". %1% = "%Operation%"
}

;Extra functionality for scripters, resume MapBasic Window runtime.
If (A_ScriptName ~= "MBWnd") {
	WinClose, WinTitle, WinText
}
ExitApp

MapInfo_RegexMatch(TableName,InputColumn,OutputColumn,Needle,MatchNum,Count){
	;Get input data
	data := MapInfo_GetData(TableName,InputColumn)

	;Regex Match
	Array_RegexMatch(data, Needle, MatchNum, Count)

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

Array_RegexMatch(data, Needle, MatchNum, Count){
	;Add "O" to Needle
	Needle := Regex_Objectify(Needle)

	;Loop through all element in array
	Loop, % data.Length()
	{
		i := 1
		While i <> 0 {
			i := RegexMatch(data[A_Index],Needle, oMatch,i)
			i := oMatch.Pos(0) + oMatch.Len(0)

			If (curCount = Count) {
				;If MatchNum = "*" return all extracted subpatterns
				If (MatchNum = "*"){
					datapart := Regex_RetAllSubPatterns(oMatch)
				} else {
					datapart := oMatch.Value(MatchNum -1)
				}
				GoTo, NextArrayElement
			} else if (Count = "*") {
				If (MatchNum = "*"){
					datapart := datapart "|" Regex_RetAllSubPatterns(oMatch)
				} else {
					datapart := datapart "|" oMatch.Value(MatchNum -1)
				}
			}
		}
	NextArrayElement:
		data[A_Index] = datapart
	}
}

;******************************************************;
;***************** HELPER FUNCTIONS *******************;
;******************************************************;

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
	iNumRows := MI.Eval("TableInfo(" TableName ",8)")
	Loop, %iNumRows%
	{
		thisData := data[A_Index -1]
		MI.Do("Update " TableName " set " ColumnName " = """ thisData """ where rowid = " A_Index)
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

