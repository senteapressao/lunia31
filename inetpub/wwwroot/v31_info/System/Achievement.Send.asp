<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request	: (POST)
	'			characterSerial,achievementHash
	'			characterName,sender,title,message,gameMoney,cashStampHash,status
	'			itemHash,stackedCount,itemSerial,instance,itemExpire
	' response	: 0,mailId
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	
	Dim rowData
	rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)
	
	' preparing request parameters
	Dim characterSerial,achievementHash
	Dim characterName,sender,title,message,gameMoney,cashStampHash,status
	Dim itemHash,stackedCount,itemSerial,instance,itemExpire
	Dim mailId,items,item
	If UBound(parameters)<8 Then Call Error("not enough parameter")	
	characterSerial	= Parameters(0)
	achievementHash	= Parameters(1)
	characterName	= Parameters(2)
	sender			= Parameters(3)
	title			= Parameters(4)
	message			= Parameters(5)
	gameMoney		= Parameters(6)
	cashStampHash	= Parameters(7)
	status			= Parameters(8)
	ReDim items(Ubound(parameters)-9)
	j=0
	For i=9 To Ubound(parameters)
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
		.SPName = "dbo.achievementSend"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterSerial",adBigInt,adParamInput,,characterSerial)
		Call .AppendParam("@achievementHash",adBigInt,adParamInput,,achievementHash)
		Call .AppendParam("@isSend",adTinyInt,adParamInput,,1)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		If ret<>0 Then
			set sphn = Nothing
			Call Error(ret+10)
		End If
	End If
	set sphn = Nothing

	On Error Resume Next

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
		ret = 0
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Rollback(characterName,achievementHash)
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

	If Err.number<>0 Then
		Call Rollback(characterName,achievementHash)
		Call Error(ret+20)
	End If

	On Error Goto 0
	
	Call Ok("0"& SEPARATOR & mailId)
%>
<!--#include file="../SendMail.asp"-->
<%
	Sub Rollback(characterName,achievementHash)
		Dim sphn,blsResult : blsResult = False

		Set sphn = new SPHelper_NoTran
		with sphn
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = characterDBconnectionString
			.SPName = "dbo.achievementSend"
			Call .InitCommand()
			Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
			Call .AppendParam("@characterSerial",adBigInt,adParamInput,,characterSerial)
			Call .AppendParam("@achievementHash",adBigInt,adParamInput,,achievementHash)
			Call .AppendParam("@isSend",adTinyInt,adParamInput,,0)
			blsResult = .ExecNoRecords()
		End with
		Set sphn = Nothing
	End Sub
%>