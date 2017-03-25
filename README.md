# MapInfo Regex Engine
This engine allows users to use Regular Expressions to manipulate MapInfo tables and datasets. The engine is designed for batch updates using RegEx. With a simple command, e.g.:

```
MI_RegexEngine.exe "Format" "TableName" "InputColumn" "OutputColumn" "ID:(\d+)" "i" "$1"
```

You can quickly use RegEx to extract data from large databases. Features include:

* Return all or a number of matches from a pattern found in a string.
* Replacing a pattern in a string with another pattern
* Formatting a pattern in a string with another pattern
* Counting the number of repititions of a given pattern in a string.
* Returning the position of a pattern in a string

The syntax has been made to try to limit the number of calls required to the engine itself. If something isn't possible without a number of calls to the engine feel free to raise an issue about it. Either there may be a better way to do what you are trying to do, or a feature may be required to do it.

The engine has also been designed with scripting in mind:

To script in the MapBasic Window Scripting simply run `note "Waiting for Regex response..."` after you make the call to the engine. For example:

```
run program """<ENGINE-PATH>\MI_RegexEngine.exe"" ""Alias"" ""Params..."""
note "Waiting for Regex response..."
note "Regex Complete! Execute more code here!"
```

The code above will run the RegEx engine, freeze the MapBasic script while the regex engine is running, and unfreeze the running script again after the RegEx engine has completed it's task.

Alternatively if you are writing code in MapBasic to be compiled to MBX you can use ShellAndWait() to run the engine synchronously with your MBX. See: http://stackoverflow.com/a/41250939/6302131 for more information regarding this.

Finally if the output column name is `$msg` or `$message` then all outputs from the Regex Engine will be outputted to the message window. This can be extremely useful for experimentation and testing!

Note there is currently a bug with this functionality:

If blank lines are found in the column currently there is a bug where these data entries are not printed to the message window. However when updating the MapInfo table everything will work as normal.

# TODO:
  Add default GUI.
  Add a way to select HWND to perform regex on. Perhaps in tablename:: @{myHwnd}|MyTableName
  Ability to use multiple input columns e.g.: `"Col1 & Col2 & Col3"`

# SYNTAX OVERVIEW:

## Regex Match

### Description:

Regex Match reads all data from the column `InputColumn` of table `TableName`, uses RegEx to find the needle specified, and returns the found data to the column `OutputColumn` of table `TableName`. This is the only function that can return all matches

### Code:

```
MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "MatchNum" "Count"
```

**Alias** can be anything that matches with the regex `i)(?:(regex)?\s*match|m)`. E.G. "Regex Match", "RegexMatch", "Match", "M"
**MatchNum** can be any positive integer or `*` representing 'all matches'.
**Count** can be any positive integer or `*` representing 'all matches'.

If Count = 1 then the 1st match will be returned

If Count = 2 then the 2nd match will be returned

If Count = * then all matches will be returned seperated by a |

If MatchNum = 0 then the whole pattern is returned

If MatchNum = 1 then the 1st captured subpatterns is returned

If MatchNum = 2 then the 2nd captured subpatterns is returned

If MatchNum = * then the whole pattern and all captured subpatterns are returned seperated by ";"


##  Regex Replace:

### Description:

Regex Replace reads all data from the column `InputColumn` of table `TableName`, uses RegEx to find the needle specified, replaces this needle given the replacement specified, and returns the found data to the column `OutputColumn` of table `TableName`.

### Code:

```
MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "Replacement"
```

**Alias** can be anything that matches with the regex `i)(?:(regex)?\s*replace|r)`. E.G. `Regex Replace`, `Replace` , `R`

**Replacement** is the string to be substituted for each match, which is plain text (not a regular expression). It may include backreferences like $1, which brings in the substring from Haystack that matched the first subpattern. The simplest backreferences are $0 through $9, where $0 is the substring that matched the entire pattern, $1 is the substring that matched the first subpattern, $2 is the second, and so on. For backreferences above 9 (and optionally those below 9), enclose the number in braces; e.g. ${10}, ${11}, and so on. For named subpatterns, enclose the name in braces; e.g. ${SubpatternName}. To specify a literal $, use $$ (this is the only character that needs such special treatment; backslashes are never needed to escape anything).

To convert the case of a subpattern, follow the $ with one of the following characters: U or u (uppercase), L or l (lowercase), T or t (title case, in which the first letter of each word is capitalized but all others are made lowercase). For example, both $U1 and $U{1} transcribe an uppercase version of the first subpattern.

Nonexistent backreferences and those that did not match anything in Haystack -- such as one of the subpatterns in "(abc)|(xyz)" -- are transcribed as empty strings.

##  Regex Format:

### Description:

Regex format takes the first match found given by needles and reformats it with the pattern specified by sFormat. sFormat follows the same rules as "Replacement" of RegexReplace above.

### Code:

```
MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "sFormat"
```

**Alias** can be anything that matches with the regex `i)(?:(regex)?\s*format|f)`. E.G. "Regex Format", "RegexFormat", "Format", "F"

##  Regex Count:

### Description:

Regex Count counts and returns the number of occurrences of a pattern.


### Code
```
MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options"
```
**Alias** can be anything that matches with the regex `(?:(regex)?\s*(get\s*)?c(ou)?nt|c)`. E.G. "Regex Count", "Count", "Cnt", "C"

##  Regex Position:

### Description:

Regex Position returns the Nth position of a pattern.

### Code:

```
MI_RegexEngine.exe "Alias" "TableName" "InputColumn" "OutputColumn" "Needle" "Options" "N"
```

**Alias** can be anything that matches with the regex `(?:(regex)?\s*(get\s*)?pos(ition)?|p)`. E.G. "Regex Position", "Position", "Pos", "P"

**N** - Default = 1

## Regex Options:

  See table of: https://autohotkey.com/docs/misc/RegEx-QuickRef.htm#Options
