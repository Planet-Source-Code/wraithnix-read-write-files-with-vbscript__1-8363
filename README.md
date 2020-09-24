<div align="center">

## Read/Write Files with VBScript


</div>

### Description

Read and write textfiles with VBScript.
 
### More Info
 
In the source.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Wraithnix](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/wraithnix.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB Script
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/wraithnix-read-write-files-with-vbscript__1-8363/archive/master.zip)





### Source Code

```
rem This function will write a file.
rem Usage: WriteStuff <filename w/path>,<text to write>
Function WriteStuff(fileout,textout)
Dim filesys,filetxt
Set filesys = CreateObject("Scripting.FileSystemObject")
Set filetxt = filesys.CreateTextFile(fileout,True)
filetxt.WriteLine(textout)
filetxt.Close
End Function
rem This function will read the contents of a textfile.
rem Usage: buffer = ReadStuff(<filename w/ path>)
Function ReadStuff(fileout)
Dim filesys,filetxt
Set filesys = CreateObject("Scripting.FileSystemObject")
Set filetxt = filesys.OpenTextFile(fileout)
ReadStuff = filetxt.ReadAll
filetxt.Close
End Function
```

