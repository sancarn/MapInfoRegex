; Taken from Exo codebase: https://github.com/Aurelain/Exo/blob/master/Exo.ahk

/**
 * Wrapper for SKAN's function (see below)
 */
getArgs(){
	CmdLine := DllCall( "GetCommandLine", "Str" )
	CmdLine := RegExReplace(CmdLine, " /ErrorStdOut", "")
	Skip    := ( A_IsCompiled ? 1 : 2 )
	argv    := Args( CmdLine, Skip )
	return argv
}


/**
 * By SKAN,  http://goo.gl/JfMNpN,  CD:23/Aug/2014 | MD:24/Aug/2014
 */
Args( CmdLine := "", Skip := 0 ) {    
	Local pArgs := 0, nArgs := 0, A := []
	pArgs := DllCall( "Shell32\CommandLineToArgvW", "WStr",CmdLine, "PtrP",nArgs, "Ptr" )
	Loop % ( nArgs )
		If ( A_Index > Skip )
			A[ A_Index - Skip ] := StrGet( NumGet( ( A_Index - 1 ) * A_PtrSize + pArgs ), "UTF-16" )  
	Return A,   A[0] := nArgs - Skip,   DllCall( "LocalFree", "Ptr", pArgs )  
}
