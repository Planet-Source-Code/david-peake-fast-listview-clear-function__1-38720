<div align="center">

## Fast ListView Clear Function


</div>

### Description

The ListView's clear method becomes slow on large lists. This function removes items faster.
 
### More Info
 
ListView Control

You've got VB!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[David Peake](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/david-peake.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/david-peake-fast-listview-clear-function__1-38720/archive/master.zip)





### Source Code

```
Public Sub ListView_Clear(lstListName As ListView)
Dim lCount As Long
Dim lLoop As Long
' Count items in listview
lCount = lstListName.ListItems.Count
' clear would probably be faster on a low number!
If lCount > 10 Then
  ' loop through (backwards) to remove items
  ' They're not visible so it's becomes fatser!!
  For lLoop = lCount To 1 Step -1
    lstListName.ListItems.Remove lLoop
  Next
Else
  lstListName.ListItems.Clear
End If
End Sub
```

