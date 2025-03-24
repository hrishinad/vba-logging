# VBA Logging

Simple, scalable class module for VBA to log messages into a textfile.

## Usage
1. Import the `clsLogFile.cls` to your Excel VBA Project
2. Use the following code in a module to get started.
```
Dim lg As clsLogFile
Set lg = New clsLogFile

lg.createFile "{Filename}.txt"
' OR
lg.openFile "{Filepath}"
```

3. Use the following code to add log messages
```
lg.log {Type As LogType (Enum)}, {txt -> log text}, {message -> Prefix for custom logtype}, {endLine = True/False}
```
> Enum used:
```
Enum LogType
  Debug_ : "[ ]: "
  Info_ : "[?]: "
  Warn_ : "[!]: "
  Error_: "[X]: "
  Custom_ : "[>]: {message}: "
End Enum
```
4. Use the code below to close the file
```
lg.closeFile
Set lg = Nothing
```

## Future
1. More filetypes
2. Optional counter
3. Timer functions
