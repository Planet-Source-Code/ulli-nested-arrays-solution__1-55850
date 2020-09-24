<div align="center">

## Nested Arrays Solution


</div>

### Description

Example
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ULLI](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ulli.md)
**Level**          |Beginner
**User Rating**    |4.6 (23 globes from 5 users)
**Compatibility**  |VB 6\.0
**Category**       |[Data Structures](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/data-structures__1-33.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ulli-nested-arrays-solution__1-55850/archive/master.zip)





### Source Code

```
Option Explicit
Private Type tNested
  Blah      As String
  Blabla     As Long
End Type
Private Type tSomething
  Stuff      As String
  MoreStuff    As Long
  aNested()    As tNested
  Whatever    As Integer
End Type
Private aSomething() As tSomething
Private Sub Form_Click()
 Dim i As Long
  ReDim aSomething(1 To 500)
  For i = 1 To 500
    ReDim aSomething(i).aNested(1 To 10)
  Next i
  aSomething(12).aNested(7).Blah = "Hallo Michael"
  Debug.Print aSomething(12).aNested(7).Blah
End Sub
```

