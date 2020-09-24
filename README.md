<div align="center">

## DAOTableToHTML


</div>

### Description

Convert a database table to HTML
 
### More Info
 
DAO Recordset Object

A reference to DAO x.x Object Library

String


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Vincent Zack](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vincent-zack.md)
**Level**          |Unknown
**User Rating**    |5.0 (5 globes from 1 user)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vincent-zack-daotabletohtml__1-1645/archive/master.zip)





### Source Code

```
Const T1 = vbTab
Const T2 = T1 & T1
Const TR = T1 & "<TR>"
Const TD = "<TD>"
Const TDEND = "</TD>"
Const TABLESTART = "<TABLE BORDER WIDTH=100%>"
Const TABLEEND = "</TABLE>"
Function HTMLTable(dbRecord As Recordset) As String
Dim strReturn As String
Dim Fld As Field
On Error GoTo Return_Zero
strReturn = strReturn & TABLESTART & vbCrLf
strReturn = strReturn & TR
For Each Fld In dbRecord.Fields
  strReturn = strReturn & TD & Fld.Name & TDEND
Next Fld
strReturn = strReturn & vbCrLf
dbRecord.MoveFirst
While Not dbRecord.EOF
  strReturn = strReturn & TR
  For Each Fld In dbRecord.Fields
    strReturn = strReturn & TD & Fld.Value & TDEND
  Next Fld
  strReturn = strReturn & vbCrLf
dbRecord.MoveNext
Wend
strReturn = strReturn & TABLEEND
Return_Zero:
HTMLTable = strReturn
End Function
```

