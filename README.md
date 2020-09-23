<div align="center">

## RevInStr


</div>

### Description

Visual Basic 6 has a new function called InStrRev, which searches a string from right to left. I found it very usefeul, so much so that a project of mine relies completely on it. When I tried to work on the project at another location on VB5 I found that the function did not exist. So, I wrote it.

I left out the compare method, you can add it if you want
 
### More Info
 
RevInStr(String1 As String, String2 As String)



Example:

let positon = RevInStr("Http://www.mypage.com/","/")

In this case position would be 22, not 1

The Integer returned is the postion of the String 2 in String 1.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Shane Presley](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shane-presley.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/shane-presley-revinstr__1-1534/archive/master.zip)

### API Declarations

	None


### Source Code

```
Public Function RevInStr(String1 As String, String2 As String) As Integer
  Dim pos As Integer
  Dim pos2 As Integer
  Let pos2 = Len(String1)
  Do
    Let pos = (InStr(pos2, String1, String2))
    Let pos2 = pos2 - 1
  Loop Until pos > 0 Or pos2 = 0
  Let RevInStr = pos
End Function
```

