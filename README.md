<div align="center">

## 95/NT username


</div>

### Description

95/NT username

"Joseph P. Fisher" <jfisher@cellone.net>
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Newsgroup Posting](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/newsgroup-posting.md)
**Level**          |Unknown
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/newsgroup-posting-95-nt-username__1-521/archive/master.zip)

### API Declarations

```
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal
lpBuffer As String, nSize As Long) As Long
```


### Source Code

```
gsUserId = ClipNull(GetUser())
Function GetUser() As String
  Dim lpUserID As String
  Dim nBuffer As Long
  Dim Ret As Long
  lpUserID = String(25, 0)
  nBuffer = 25
  Ret = GetUserName(lpUserID, nBuffer)
  If Ret Then
  GetUser$ = lpUserID$
  End If
End Function
Function ClipNull(InString As String) As String
  Dim intpos As Integer
  If Len(InString) Then
   intpos = InStr(InString, vbNullChar)
   If intpos > 0 Then
    ClipNull = Left(InString, intpos - 1)
   Else
    ClipNull = InString
   End If
  End If
End Function
```

