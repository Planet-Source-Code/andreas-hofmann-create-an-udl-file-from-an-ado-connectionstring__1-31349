<div align="center">

## Create an UDL\-File from an ADO\-ConnectionString


</div>

### Description

Creates an UDL-File (Universal Data Link) from an existing ADO-ConnectionString. Unfortunately the DataLinks-Object in the "Microsoft OLE DB Service Component 1.0 Type Library" provides some prompting dialogs to choose an ADO-Connection, but a save method to get an UDL-file is missing. So I coded this. Note that this wasn't as easy as it looks now, cause an UDL-file is no normal INI-file, although it seems to be one. The first thing is, that it must be saved in unicode. The second and very astonishing thing is, that the second line, which seems to be a missable comment, is very important and must be exactly as it is. Otherwise the UDL-file wont work! Comments and votes are welcome.
 
### More Info
 
ConnectionString As String, Filename As String

Needs to have a reference to the "Microsoft Scripting Runtime".

None, but creates as file.

Overwrites the given Filename, if already exists.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Andreas Hofmann](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/andreas-hofmann.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/andreas-hofmann-create-an-udl-file-from-an-ado-connectionstring__1-31349/archive/master.zip)





### Source Code

```
Public Sub CreateUDLFile(ConnectionString As String, FileName As String)
 Dim FSO As New Scripting.FileSystemObject
 Dim TXT As Scripting.TextStream
 ' Create a File in Unicode-Mode
 Set TXT = FSO.CreateTextFile(FileName, True, True)
 With TXT
 .WriteLine "[oledb]"
 ' This line needs to be exactly as it is
 .WriteLine "; Everything after this line is an OLE DB initstring"
 .WriteLine ConnectionString
 .Close
 End With
End Sub
```

