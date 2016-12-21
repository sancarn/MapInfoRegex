Operation = %1%
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
	Msgbox, Command not found: "%1%"
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
	ExitApp
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