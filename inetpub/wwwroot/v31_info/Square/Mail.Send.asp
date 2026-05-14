<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : (POST)
	'			characterName,sender,title,message,gameMoney,cashStampHash,status,mailHash
	'			{itemHash,stackedCount,itemSerial,instance,itemExpire}
	' response : 0,mailId
	'			characterName,sender,title,message,gameMoney,cashStampHash,status,mailHash
	'			{itemHash,stackedCount,itemSerial,instance,itemExpire}
	
	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	
	Dim rowData
	rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)
	
	' preparing request parameters
	Dim characterName,sender,title,message,gameMoney,cashStampHash,status,mailHash
	Dim itemHash,stackedCount,itemSerial,instance,itemExpire
	Dim mailId,items,item

	If UBound(Parameters)<7 Then Call Error("not enough parameter")
	characterName	= Parameters(0)
	sender			= Parameters(1)
	title			= Parameters(2)
	message			= Parameters(3)
	gameMoney		= Parameters(4)
	cashStampHash	= Parameters(5)
	status			= Parameters(6)
	mailHash		= Parameters(7)
	ReDim items(Ubound(parameters)-8)
	j=0
	For i=8 To Ubound(parameters)
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
		.SPName = "dbo.SendMail"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@sender",adVarWChar,adParamInput,50,sender)
		Call .AppendParam("@title",adVarWChar,adParamInput,20,title)
		Call .AppendParam("@message",adVarWChar,adParamInput,150,message)
		Call .AppendParam("@gameMoney",adBigInt,adParamInput,,gameMoney)
		Call .AppendParam("@cashStampHash",adInteger,adParamInput,,cashStampHash)
		Call .AppendParam("@status",adUnsignedTinyInt,adParamInput,,status)
		Call .AppendParam("@mailId",adBigInt,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		mailId = sphn.GetParamValue("@mailId")
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
			Call .AppendParam("@stackedCount",adSmallInt,adParamInput,,stackedCount)
			Call .AppendParam("@itemSerial",adBigInt,adParamInput,,itemSerial)
			Call .AppendParam("@instance",adBigInt,adParamInput,,instance)
			Call .AppendParam("@itemExpire",adDBTimeStamp,adParamInput,,FormatDt(itemExpire,"SQL_TM"))
			blsResult = .ExecNoRecords()
		End with
		set sphn = Nothing
	Next
	
	Call Ok("0"& SEPARATOR & mailId & SEPARATOR & rowData)
%>