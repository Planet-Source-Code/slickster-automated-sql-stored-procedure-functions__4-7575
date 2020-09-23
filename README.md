<div align="center">

## Automated SQL Stored procedure functions


</div>

### Description

I was tired of writing code to execute store procedures so I wrote these functions that do most everything for me. Just supply the stored procedure name and an array of parameter values. Also provide a recordset or return value variable depending on what function you are using. SAVES ALOT OF TIME! There are some examples subs at the bottom...
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Slickster](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/slickster.md)
**Level**          |Advanced
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |ASP \(Active Server Pages\)
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__4-5.md)
**World**          |[ASP / VbScript](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/asp-vbscript.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/slickster-automated-sql-stored-procedure-functions__4-7575/archive/master.zip)





### Source Code

```
'--------------------------------------------Start Function getRS----------------------------------------------------------
'This function is used to return a recordset
Function getRS(strSPName, aParamaters(), byRef rsNew)
	'on error resume next
	dim strStoredProcedureName
	strStoredProcedureName = strSPName
	dim cmdGetRS
	set cmdGetRS = Server.CreateObject("ADODB.Command")
	dim rsGetRS
	set rsGetRS = Server.CreateObject("ADODB.Recordset")
	dim connNewConnection
	connNewConnection = GetOpenConnection
	cmdGetRS.ActiveConnection = connNewConnection
	cmdGetRS.CommandType = adCmdStoredProc
	cmdGetRS.CommandText = strStoredProcedureName
	rsGetRS.CursorType = adOpenStatic
	rsGetRS.CursorLocation = adUseClient
	rsGetRS.LockType = adLockReadOnly
	'Parameter object to split up the parameter collection object
	dim param
	'Counter to Sync parameter array values with stored procedure attributes
	dim intCount
	intCount = 0
	'Loop through parameter collection
	for each param in cmdGetRS.Parameters
		'Skip the Return value
		if param.name <> "RETURN_VALUE" then
			Param.value = GetDataTypeEnum(param.Type,aParamaters(intCount))
			intCount = intCount + 1
		end if
	next
	'Open a recordset with the results
	rsGetRS.Open cmdGetRS
	'Set the recordset to be returned
	set rsNew = rsGetRS.Clone
	rsGetRS.close
	set rsGetRS = nothing
	set cmdGetRS = nothing
	if err.number <> 0 then
		getRS = false
	else
		getRS = true
	end if
End Function
'--------------------------------------------End Function getRS------------------------------------------------------------
'--------------------------------------------Start Function addRS----------------------------------------------------------
'This Function add an item to the database and will return a value if the stored procedure supplies one
Function addRS(strSPName,params(),byRef strOutputParam)
	'On Error resume next
	dim strStoredProcedureName
	strStoredProcedureName = strSPName
	Dim cmdAddRS
	Set cmdAddRS = Server.CreateObject("ADODB.Command")
	dim connNewConnection
	'I have an external function to return a db connection. Just use a dsn or a connection string
	connNewConnection = GetOpenConnection
	cmdAddRS.ActiveConnection = connNewConnection
	cmdAddRS.CommandType = adCmdStoredProc
	cmdAddRS.CommandText = strStoredProcedureName
	dim param
	dim blnOutput
	dim intCount
	intCount = 0
	blnOutput = false
	for each param in cmdAddRS.Parameters
		if param.name <> "RETURN_VALUE" then
			if (GetParameterDirectionEnum(param.Direction) = "adParamOutput") or (GetParameterDirectionEnum(param.Direction) = "adParamInputOutput") then
				'Let's the code know if there is a output value ie: Item ID
				blnOutput = true
				strOutputParam = Param.name
			else
				Param.value = GetDataTypeEnum(param.Type,params(intCount))
			end if
			intCount = intCount + 1
		end if
	next
	cmdAddRS.Execute
	'Set the return value to be returned
	if blnOutPut = True then
		strOutputParam = cmdAddRS.Parameters(strOutputParam).Value
	end if
	set cmdAddRS = nothing
	if err.number <> 0 then
		addRS = False
	else
		addRS = True
	End if
End Function
'--------------------------------------------End Function addRS------------------------------------------------------------
'--------------------------------------------Start Function updateRS-------------------------------------------------------
'This function performs an update for a particular item.
Function updateRS(strSPName,params())
	'On Error resume next
	dim strStoredProcedureName
	strStoredProcedureName = strSPName
	dim cmdUpdateRS
	set cmdUpdateRS = Server.CreateObject("ADODB.Command")
	dim rsUpdateRS
	set rsUpdateRS = Server.CreateObject("ADODB.Recordset")
	dim connNewConnection
	connNewConnection = GetOpenConnection
	cmdUpdateRS.ActiveConnection = connNewConnection
	cmdUpdateRS.CommandType = adCmdStoredProc
	cmdUpdateRS.CommandText = strStoredProcedureName
	dim param
	dim intCount
	dim blnOutPut
	intCount = 0
	for each param in cmdUpdateRS.Parameters
		if param.name <> "RETURN_VALUE" then
			Param.value = GetDataTypeEnum(param.Type,params(intCount))
			intCount = intCount + 1
		end if
	next
	cmdUpdateRS.Execute
	if blnOutPut = True then
		strOutputParam = cmdUpdateRS.Parameters(strOutputParam).Value
	end if
	set cmdUpdateRS = nothing
	if err.number <> 0 then
		updateRS = False
	else
		updateRS = True
	End if
End Function
'--------------------------------------------End Function updateRS---------------------------------------------------------
'--------------------------------------------Start Function GetParameterDirectionEnum--------------------------------------
'This function determines the direction of the parameter
Function GetParameterDirectionEnum(lngDirection)
  Select Case lngDirection
    Case 0 'adParamUnknown
      GetParameterDirectionEnum = "adParamUnknown"
    Case 1 'adParamInput
      GetParameterDirectionEnum = "adParamInput"
    Case 2 'adParamOutput
      GetParameterDirectionEnum = "adParamOutput"
    Case 3 'adParamInputOutput
      GetParameterDirectionEnum = "adParamInputOutput"
    Case 4 'adParamReturnValue
      GetParameterDirectionEnum = "adParamReturnValue"
    Case Else
						GetParameterDirectionEnum = "<B>Direction Not Found</B>"
  End Select
End Function
'--------------------------------------------End Function GetParameterDirectionEnum----------------------------------------
'--------------------------------------------Start Function GetDataTypeEnum------------------------------------------------
'This function is used to determine the parameter type and validates the data accordingly.
Function GetDataTypeEnum(lngType,ByRef value)
  Select Case lngType
    Case 0
      GetDataTypeEnum = "adEmpty"
    Case 2
      GetDataTypeEnum = "adSmallInt"
    Case 3
      GetDataTypeEnum = CLng(value)
    Case 4
      GetDataTypeEnum = "adSingle"
    Case 5
      GetDataTypeEnum = CDBL(value)
    Case 6
      GetDataTypeEnum = CCur(value)
    Case 7
      GetDataTypeEnum = Cdate(value)
    Case 8
      GetDataTypeEnum = CStr(value)
    Case 9
      GetDataTypeEnum = "adIDispatch"
    Case 10
      GetDataTypeEnum = "adError"
    Case 11
      GetDataTypeEnum = CBool(value)
    Case 12
      GetDataTypeEnum = "adVariant"
    Case 13
      GetDataTypeEnum = "adIUnknown"
    Case 14
      GetDataTypeEnum = CDBL(value)
    Case 16
      GetDataTypeEnum = "adTinyInt"
    Case 17
      GetDataTypeEnum = "adUnsignedTinyInt"
    Case 18
      GetDataTypeEnum = "adUnsignedSmallInt"
    Case 19
      GetDataTypeEnum = "adUnsignedInt"
    Case 20
      GetDataTypeEnum = "adBigInt"
    Case 21
      GetDataTypeEnum = "adUnsignedBigInt"
    Case 64
      GetDataTypeEnum = "adFileTime"
    Case 72
      GetDataTypeEnum = "adGUID"
    Case 128
      GetDataTypeEnum = "adBinary"
    Case 129
      GetDataTypeEnum = "adChar"
    Case 130
      GetDataTypeEnum = "adWChar"
    Case 131
      GetDataTypeEnum = "adNumeric"
    Case 132
      GetDataTypeEnum = "adUserDefined"
    Case 133
      GetDataTypeEnum = "adDBDate"
    Case 134
      GetDataTypeEnum = CDate(value)
    Case 135
      GetDataTypeEnum = CDate(value)
    Case 136
      GetDataTypeEnum = "adChapter"
    Case 138
      GetDataTypeEnum = "adPropVariant"
    Case 139
      GetDataTypeEnum = "adVarNumeric"
    Case 200
      GetDataTypeEnum = CStr(value)
    Case 201
      GetDataTypeEnum = "adLongVarChar"
    Case 202
      GetDataTypeEnum = "adVarWChar"
    Case 203
      GetDataTypeEnum = "adLongVarWChar"
    Case 204
      GetDataTypeEnum = "adVarBinary"
    Case 205
      GetDataTypeEnum = "adLongVarBinary"
    Case 8192
      GetDataTypeEnum = "adArray"
    Case Else
      'GetDataTypeEnum = "<B>Type Constant Not Found</B>"
  End Select
End Function
'--------------------------------------------End Function GetDataTypeEnum--------------------------------------------------
'The following are example procedures to implement the preceding functions.
'Examplegetlist
'ExampleaddCountry
Sub ExampleAddCountry
	dim params(3)
	params(0) = "0"
	params(1) = "Test" & Now
	params(2) = "0"
	dim blnSucceeded
	dim strOutput
	blnSucceeded = addRS("sp_insert_c_Country",params,strOutput)
	if blnSucceeded = True then
		getlist strOutput
		dim uparams(3)
		uparams(0) = strOutput
		uparams(1) = "0"
		uparams(2) = "renamed" & now
		blnSucceeded = updateRS("sp_update_c_Country",uparams)
		if blnSucceeded = True then
			getlist strOutput
		else
			Response.Write "ERROR: Update"
		end if
	else
		Response.Write "ERROR: " & strOutput
	end if
End Sub
'This example funtion returns a list of countries or a single country(if a country ID is provided)
Sub ExampleGetList(itemID)
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	dim params(1)
	if itemID = "" then
		params(0) = "0"
	else
		params(0) = itemID
	end if
	dim blnSucceeded
	blnSucceeded = getRS("sp_select_c_Country",params,rs)
	if blnSucceeded = True then
		if rs.eof then
			Response.Write "empty"
		else
			while not rs.EOF
				Response.Write "<BR>" & rs("intCCountryIDPK") & "-" & rs("vchCCountryName")
				rs.Movenext
			wend
		End if
	else
		Response.Write "Error"
	end if
End Sub
```

