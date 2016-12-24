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



/**
 * ALTERNATIVE Extended to return an array.
 */
 
 aGetArgs(bAll := false) {
	/*
	Based on code By SKAN,  http://goo.gl/JfMNpN,  CD:23/Aug/2014 | MD:24/Aug/2014
	Modified by Dougal 17Jan15
	Returns array of strings, element 0 contains array count
	Includes commandline parts if bAll = true
	Otherwise strips non-parameter parts depending on how it was called:
		Compiled script
			removes compiled script name
		Script
			removes AutoHotkey executable and script name
		SciTe4AutoHotkey
			removes AutoHotkey executable, /ErrorStdOut parameter (if present) and script name
	*/
	sCmdLine := DllCall( "GetCommandLine", "Str" )
	pArgs := 0, nArgs := 0, aArgs := []
	pArgs := DllCall( "Shell32\CommandLineToArgvW", "WStr",sCmdLine, "PtrP",nArgs, "Ptr" )
	; get command line parts from memory
	Loop % (nArgs)
			aArgs.insert(StrGet(NumGet((A_Index - 1 ) * A_PtrSize + pArgs ), "UTF-16"))
	if not bAll { ; remove calling program parts
		; remove AutoHotkey.exe if not compiled script
		if not A_IsCompiled {
			aArgs.remove(1)
			nArgs -= 1
		}
		; remove /ErrorStdOut (run within SciTE4AutoHotkey)
		if (aArgs[1] = "/ErrorStdOut") {
			aArgs.remove(1)
			nArgs -= 1
		}
		; remove (compiled) script name
		aArgs.remove(1)
		nArgs -= 1
	}
	aArgs[0] := nArgs
	DllCall("LocalFree", "Ptr", pArgs)
	return, aArgs
}
