<div align="center">

## ASP Format Function


</div>

### Description

This function operates similarly to the VB Format function with one big exception. The "#" character is used to represent any single character. You can trim all non alphanumeric characters out and reformat them to stay consistant.

Usefull for credit cards, zipcodes, phone numbers, etc...
 
### More Info
 
Format("1234567890123", "(###) ###-#### x######") would return "(123) 456-7890 x123"

Format("4111111111111111", "####-####-####-####")

would return "4111-1111-1111-1111"


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Brian Reeves](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/brian-reeves.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Validation/ Processing](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/validation-processing__4-16.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/brian-reeves-asp-format-function__4-8175/archive/master.zip)

### API Declarations

Open Source


### Source Code

```
'******
'**		Formats a string to include standard sets.
'**
'**		Example:	Format("1234567890", "(###) ###-####")
'**			Result =	(123) 456-7890
'**		Modified 01/09/03 to allow extended format mask that will
'**			not return extra ###'s brian reeves
'******
Public Function Format(sValue, sMask)
	Dim iPlaceHolder
	Dim sTempValue
	Dim sResult
	sTempValue = CStr(sValue)
	sResult = sMask
	Do Until InStr(sResult, "#") = 0
		iPlaceHolder = InStr(sResult, "#")
		sResult = Replace(sResult, "#", Left(sTempValue, 1), 1, 1)
		sTempValue = Mid(sTempValue, 2)
		If Len(sTempValue) = 0 Then sResult = Left(sResult, iPlaceHolder)
	Loop
	Format = sResult
End Function
```

