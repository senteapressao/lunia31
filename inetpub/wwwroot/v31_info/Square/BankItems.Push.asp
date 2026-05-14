<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,bagNumber,positionNumber
	'			,itemHash,stackedCount,itemSerial,instance,itemExpire
	'			,moveCount,maxStackCount
	' response : characterName,bagNumber,positionNumber
	'			,itemHash,stackedCount,itemSerial,instance,itemExpire (inven)
	'			,itemHash,stackedCount,itemSerial,instance,itemExpire (bank)
	'			,moveCount

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName,bagNumber,positionNumber _
		,itemHash,stackedCount,itemSerial,instance,itemExpire _
		,moveCount,maxStackCount
	If UBound(parameters)<9 Then Call Error("not enough parameter")
	characterName	= Parameters(0)
	bagNumber		= Parameters(1)
	positionNumber	= Parameters(2)
	itemHash		= Parameters(3)
	stackedCount	= Parameters(4)
	itemSerial		= Parameters(5)
	instance		= Parameters(6)
	itemExpire		= Parameters(7)
	moveCount		= Parameters(8)
	maxStackCount	= Parameters(9)
	
	Dim itemHash_bank,stackedCount_bank,itemSerial_bank,instance_bank,itemExpire_bank

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.BankItems_Push"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@bagNumber",adUnsignedTinyInt,adParamInput,,bagNumber)
		Call .AppendParam("@positionNumber",adUnsignedTinyInt,adParamInput,,positionNumber)
		Call .AppendParam("@moveCount",adSmallInt,adParamInput,,moveCount)
		Call .AppendParam("@maxStackCount",adSmallInt,adParamInput,,maxStackCount)
		Call .AppendParam("@itemHash",adInteger,adParamInputOutput,,itemHash)
		Call .AppendParam("@stackedCount",adSmallInt,adParamInputOutput,,stackedCount)
		Call .AppendParam("@itemSerial",adBigInt,adParamInputOutput,,itemSerial)
		Call .AppendParam("@instance",adBigInt,adParamInputOutput,,instance)
		Call .AppendParam("@itemExpire",adDBTimeStamp,adParamInputOutput,,itemExpire)
		Call .AppendParam("@itemHash_bank",adInteger,adParamOutput,,null)
		Call .AppendParam("@stackedCount_bank",adSmallInt,adParamOutput,,null)
		Call .AppendParam("@itemSerial_bank",adBigInt,adParamOutput,,null)
		Call .AppendParam("@instance_bank",adBigInt,adParamOutput,,null)
		Call .AppendParam("@itemExpire_bank",adDBTimeStamp,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		itemHash			= sphn.GetParamValue("@itemHash")
		stackedCount		= sphn.GetParamValue("@stackedCount")
		itemSerial			= sphn.GetParamValue("@itemSerial")
		instance			= sphn.GetParamValue("@instance")
		itemExpire			= sphn.GetParamValue("@itemExpire")
		itemHash_bank		= sphn.GetParamValue("@itemHash_bank")
		stackedCount_bank	= sphn.GetParamValue("@stackedCount_bank")
		itemSerial_bank		= sphn.GetParamValue("@itemSerial_bank")
		instance_bank		= sphn.GetParamValue("@instance_bank")
		itemExpire_bank		= sphn.GetParamValue("@itemExpire_bank")
	End If
	set sphn = Nothing

	retString = characterName & SEPARATOR & bagNumber & SEPARATOR & positionNumber &_
			SEPARATOR & itemHash & SEPARATOR & stackedCount &_
			SEPARATOR & itemSerial & SEPARATOR & instance & SEPARATOR & FormatDt(itemExpire,"SQL_TM") &_
			SEPARATOR & itemHash_bank & SEPARATOR & stackedCount_bank &_
			SEPARATOR & itemSerial_bank & SEPARATOR & instance_bank & SEPARATOR & FormatDt(itemExpire_bank,"SQL_TM") &_
			SEPARATOR & moveCount

	Call OK(retString)
%>