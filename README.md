# MapInfoRegex

A RegEx Engine for MapInfo

## FEATURES:
Regular Expressions support for MapInfo!
Support brought for the following functions:
  Regex Match     - Pattern found; Data to chosen column
  Regex Replace   - Pattern found; Patterns replaced; Data to chosen column
  Regex Format    - Pattern and subpatterns found; Reformated with \1 \2 \3 ...; Data to chosen column
  Regex Count		 - Pattern counted. Count returned to chosen column.
  Regex Position	 - Pattern found. Position returned to chosen column

## SYNTAX OVERVIEW:
###  Regex Match:
    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "MatchNum" "Count"
      Alias: "Regex Match", "RegexMatch", "Match", "M" - CASE INSENSITIVE
      MatchNum: 1, 2, 3, ... or *
      Count: 1, 2, 3, ... or *

    Description:
      Regex Match reads all data from the column <InputColumn> of table <TableName>,
      uses RegEx to find the needle specified, and returns the found data to the
      column <OutputColumn> of table <TableName>.

      If Count = 1 then the 1st match will be returned
      If Count = 2 then the 2nd match will be returned
      If Count = * then     all matches will be returned seperated by a |

      If MatchNum = 0 then the whole pattern is returned
      If MatchNum = 1 then the 1st captured subpatterns is returned
      If MatchNum = 2 then the 2nd captured subpatterns is returned
      If MatchNum = * then the whole pattern and all captured subpatterns
      are returned seperated by ";"

###  Regex Replace:
    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "Replacement"
      Alias: "Regex Replace", "RegexReplace", "Replace", "R" - CASE INSENSITIVE
      Replacement - The string to be substituted for each match, which is plain text (not a regular
       expression). It may include backreferences like $1, which brings in the substring from Haystack
  		  that matched the first subpattern. The simplest backreferences are $0 through $9, where $0 is
  		  the substring that matched the entire pattern, $1 is the substring that matched the first subpattern,
  		  $2 is the second, and so on. For backreferences above 9 (and optionally those below 9), enclose the
  		  number in braces; e.g. ${10}, ${11}, and so on. For named subpatterns, enclose the name in braces;
  		  e.g. ${SubpatternName}. To specify a literal $, use $$ (this is the only character that needs such
  		  special treatment; backslashes are never needed to escape anything).

       To convert the case of a subpattern, follow the $ with one of the following characters: U or u (uppercase),
       L or l (lowercase), T or t (title case, in which the first letter of each word is capitalized but all others are
       made lowercase). For example, both $U1 and $U{1} transcribe an uppercase version of the first subpattern.

       Nonexistent backreferences and those that did not match anything in Haystack -- such as one of the
       subpatterns in "(abc)|(xyz)" -- are transcribed as empty strings.

###  Regex Format:
    Description:
      Regex format takes the first match found given by needles and reformats it with the pattern specified
				by sFormat. sFormat follows the same rules as "Replacement" of RegexReplace above.
    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "sFormat"
      Alias: "Regex Format", "RegexFormat", "Format", "F" - CASE INSENSITIVE

###  Regex Count:
    Description:
      Regex Count counts and returns the number of occurrences of a pattern.
    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options"
      Alias: "Regex Count", "Count", "Cnt", "C" (OTHERS: "(?:(regex)?\s*(get\s*)?c(ou)?nt|c)")

###  Regex Position:
    Description:
      Regex Position returns the Nth position of a pattern.
    CMD: MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "N"
      Alias: "Regex Position", "Position", "Pos", "P" (OTHERS: "(?:(regex)?\s*(get\s*)?pos(ition)?|p)")
      N: Default = 1

##Regex Options:

  See table of: https://autohotkey.com/docs/misc/RegEx-QuickRef.htm#Options

##MapBasic Window Scripting:
  If you are making a script in the MapBasic window you can call the Regex Engine with this:
    UNCOMPILED:
      run program """C:\Program Files\AutoHotkey\AutoHotkey.exe"" ""C:\Users\jwa\Desktop\MapInfoRegex-master\MI_RegexEngine.ahk"" ""Alias"" ""Params..."""
      note "Waiting for Regex response..."
      note "Regex Complete! Execute more code here!"

    COMPILED:
      run program """<ENGINE-PATH>\MI_RegexEngine.exe"" ""Alias"" ""Params..."""
      note "Waiting for Regex response..."
      note "Regex Complete! Execute more code here!"

  The code above will run the RegEx engine, freeze the MapBasic script while the regex engine is running, and unfreeze
   the running script again after the RegEx engine has completed it's task.

##MBX Scripting
  If you are writing code in MapBasic to be compiled to MBX, use ShellAndWait(). (See: http://stackoverflow.com/a/41250939/6302131 for an implementation of this.)

##OUTPUT TO MESSAGE WINDOW:
  If the output column name is $msg or $message then the change will be outputted to the message window. Useful for experimentation!
  BUG:
  If blank lines are found in the column currently there is a bug where these data entries are not printed to the message window.
  However when updating the MapInfo table everything will work as normal.


TODO:
  Add default GUI.
  Add a way to select HWND to perform regex on. Perhaps in tablename:: @{myHwnd}|MyTableName
