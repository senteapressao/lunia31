<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : (POST)
	'			mailId,characterName,gameMoney
	'			itemHash,stackedCount,itemSerial,instance,itemExpire
	' response : 0

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	
	Dim rowData
	rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)
	
	' preparing request parameters
	Dim mailId,characterName,gameMoney
	Dim itemHash,stackedCount,itemSerial,instance,itemExpire
	Dim items,item

	If UBound(Parameters)<2 Then Call Error("not enough parameter")
	mailId			= Parameters(0)
	characterName	= Parameters(1)
	gameMoney		= Parameters(2)
	ReDim items(Ubound(parameters)-3)
	j=0
	For i=3 To Ubound(parameters)
		If parameters(i)<>"" Then
			items(j) = parameters(i)
			j=j+1
		End If
	Next
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.UpdateMail"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@mailId",adBigInt,adParamInput,,mailId)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@title",adVarWChar,adParamInput,20,null)
		Call .AppendParam("@message",adVarWChar,adParamInput,150,null)
		Call .AppendParam("@gameMoney",adBigInt,adParamInput,,gameMoney)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
	End If
	set sphn = Nothing
	
	For i=0 To Ubound(items)
		item = Split(items(i),SUBSEPARATOR)
		itemHash	= item(0)
		stackedCount= item(1)
		itemSerial	= item(2)
		instance	= item(3)
		itemExpire	= item(4)

		Set sphn = new SPHelper_NoTran
		with sphn
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = characterDBconnectionString
			.SPName = "dbo.SendMailItem"
			Call .InitCommand()
			Call .AppendParam("@mailId",adBigInt,adParamInput,,mailId)
			Call .AppendParam("@itemHash",adInteger,adParamInput,,itemHash)
			Call .AppendParam("@stackedCount",adInteger,adParamInput,,stackedCount)
			Call .AppendParam("@itemSerial",adBigInt,adParamInput,,itemSerial)
			Call .AppendParam("@instance",adBigInt,adParamInput,,instance)
			Call .AppendParam("@itemExpire",adDBTimeStamp,adParamInput,,FormatDt(itemExpire,"SQL_TM"))
			blsResult = .ExecNoRecords()
		End with
		set sphn = Nothing
	Next

	Call Ok("0"& SEPARATOR & mailId)
%>