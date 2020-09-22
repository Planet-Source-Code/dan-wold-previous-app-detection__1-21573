<div align="center">

## Previous App Detection


</div>

### Description

This easy code took me about 20 seconds to write, this detects on the Form_Load() Event if another Instance of the exe is already running, if it is it displays an error dialog, else it just runs the program, this is basicly for newbies, please vote for me, Happy Coding...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Dan Wold](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/dan-wold.md)
**Level**          |Beginner
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/dan-wold-previous-app-detection__1-21573/archive/master.zip)

### API Declarations

(None)


### Source Code

```
If App.PrevInstance = True Then
MsgBox "There Is Already Another Instance Of This Application Running, Please Close It And Try Again.", vbExclamation, "Error"
End
End If
```

