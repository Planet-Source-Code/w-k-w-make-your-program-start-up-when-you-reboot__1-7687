<div align="center">

## Make Your Program Start Up When You Reboot


</div>

### Description

This is to make your program start up when you reboot your computer just like other programs you see do. **Must See** Please Vote for me!
 
### More Info
 
Just put this code in the form of your program.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[WÏ©kÊÐ\_WÎ££ÿ](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/w-k-w.md)
**Level**          |Beginner
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/w-k-w-make-your-program-start-up-when-you-reboot__1-7687/archive/master.zip)





### Source Code

```
Private Sub Form_Load()
Dim strDir As String
Dim intDir As Integer
Dim strpath As String
intDir = Len(CurDir()) 'Gets length of Current directory
'Gets last character from CurDir() and checks if it's a \
strDir = Mid(CurDir(), intDir)
If strDir = "\" Then
 strpath = CurDir() & App.EXEName & ".exe"
 'If is in main drive like C:\ or D:\ it simply
 'puts the file name, "C:\Blah.exe"
Else
 strpath = CurDir() & "\" & App.EXEName & ".exe"
 'If CurDir() returns no \ then its in a folder
 'and will necessitate a \ inserted so that it looks like
 'C:\Folder\Blah.exe and NOT like C:\FolderBlah.exe
End If
On Error GoTo Death
'Error statement allows this code to run if it is already
'in the Start Menu
FileCopy strpath, _
"C:\WINDOWS\Start Menu\Programs\StartUp\" _
& App.EXEName & ".exe"
Death:
Exit Sub
Resume Next
End Sub
```

