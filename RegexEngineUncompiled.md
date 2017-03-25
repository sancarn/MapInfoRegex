# How to run the regex engine while uncompiled .ahk

```
cd C:\Program Files\AutoHotkey\
AutoHotkey.exe "C:\path\to\regex\engine\MI_RegexEngine.ahk" "Alias" "Params..."
```

Or in mapinfo:

```
run program """C:\Program Files\AutoHotkey\AutoHotkey.exe"" ""C:\path\to\regex\engine\MI_RegexEngine.ahk"" ""Alias"" ""Params..."""
note "Waiting for Regex response..."
```
