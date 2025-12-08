Set WshShell = CreateObject("WScript.Shell") 

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")

' system path to 7zip
szpath = """C:\Program Files\7-Zip\7z.exe"""

' Get current folder name to use as output filename
pathSplit = Split(Cstr(WScript.ScriptFullName), "\")
currFolder = pathSplit(UBound(pathSplit) - 1)

' create filepath that would match already existing zip file
pathSplit(UBound(pathSplit)) = currFolder
fpath = Join(pathSplit,"\") & ".zip"

' delete existing zip if exists
If fso.FileExists( fpath ) = True Then
   fso.DeleteFile( fpath )
End If


' create new zip with modinfo.json and everything in ./data
command = szpath & " a -tzip " & Chr(34) & currFolder & Chr(34) & ".zip" &" modinfo.json -ir!data\*"

WshShell.Exec command
