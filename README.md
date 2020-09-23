<div align="center">

## Quick String Parsing of a Comma Delimited String


</div>

### Description

This example will show you how parse a list a variables quickly with minimal coding.
 
### More Info
 
A string of variables separated by a common like "stuff1,stuff2,stuff3"

Returns individual variables


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mark Collins](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-collins.md)
**Level**          |Beginner
**User Rating**    |3.7 (22 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) 
**Category**       |[String Manipulation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/string-manipulation__1-5.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mark-collins-quick-string-parsing-of-a-comma-delimited-string__1-33915/archive/master.zip)





### Source Code

```
Sub TestMe()
  'Run this sub to demonstate how the parsing functions work.
  'This is an example to show begining vb programmers how to parse strings
  'quickly with out having to use mid/instr/right/left every time then need
  'the next item in a string of items.
  'Declare our working strings
  Dim myString As String
  Dim curShiz As String, curVar As String
  'Declare our name containers
  Dim Name1 As String
  Dim Name2 As String, Name3 As String, Name4 As String
  myString = "Tom,Debbie,Mark,Joanie"
  'Get tom
  curVar = GetFirstVar(myString)
  myString = GetRestOfVars(myString)
  Name1 = curVar
  'Get Debbie
  curVar = GetFirstVar(myString)
  myString = GetRestOfVars(myString)
  Name2 = curVar
  'Get MArk
  curVar = GetFirstVar(myString)
  myString = GetRestOfVars(myString)
  Name3 = curVar
  'Get joanie
  curVar = GetFirstVar(myString)
  myString = GetRestOfVars(myString)
  Name4 = curVar
  MsgBox "Name1 = " & Name1
  MsgBox "Name2 = " & Name2
  MsgBox "Name3 = " & Name3
  MsgBox "Name4 = " & Name4
End Sub
Public Function GetFirstVar(sStr As String)
  'Given a string like: "choad,flap,blah"
  'this returns "choad"
  f = InStr(1, sStr, ",")
  If f = 0 Then
    GetFirstVar = sStr
  Else
    GetFirstVar = Left(sStr, f - 1)
  End If
End Function
Public Function GetRestOfVars(sStr As String)
  'Given a string like: "choad,flap,blah"
  'this returns "flap,blah"
  f = InStr(1, sStr, ",")
  GetRestOfVars = Right(sStr, Len(sStr) - f)
End Function
```

