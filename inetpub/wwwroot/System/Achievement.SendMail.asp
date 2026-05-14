<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init() 
	' request	: (POST)
	'			characterSerial,gold
	'			itemHash,stackedCount,itemSerial,instance,itemExpire
	' response	: 0,mailId
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	
	Dim rowData
	rowData = getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)
	
	' preparing request parameters
	Dim characterSerial,achievementHash,gameMoney
	Dim mailId,items,item
	If UBound(parameters)<3 Then Call Error("not enough parameter")	
	characterSerial	= Parameters(0)
	achievementHash = Parameters(1)
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
		.SPName = "dbo.achievement_SendMail"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterSerial",adBigInt,adParamInput,,characterSerial)
		Call .AppendParam("@achievementHash",adBigInt,adParamInput,,achievementHash)
		Call .AppendParam("@gameMoney",adBigInt,adParamInput,,gameMoney)
		Call .AppendParam("@mailId",adBigInt,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret= 0
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		mailId = sphn.GetParamValue("@mailId")
	End If
	set sphn = Nothing
	Dim itemHash,stackedCount,instance,itemExpire
	For i=0 To Ubound(items)
		item = Split(items(i),SUBSEPARATOR)
		itemHash	= item(0)
		stackedCount= item(1)
		instance	= item(2)
		itemExpire	= item(3)

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
			Call .AppendParam("@itemSerial",adBigInt,adParamInput,,0)
			Call .AppendParam("@instance",adBigInt,adParamInput,,instance)
			Call .AppendParam("@itemExpire",adDBTimeStamp,adParamInput,,FormatDt(itemExpire,"SQL_TM"))
			blsResult = .ExecNoRecords()
		End with
		set sphn = Nothing
	Next

	If Err.number<>0 Then
		Call Error(ret+20)
	End If
	
	Call Ok("0"& SEPARATOR & mailId)
%>