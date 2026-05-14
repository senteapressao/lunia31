<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	'			,bagNumber_source,positionNumber_source
	'			,bagNumber_target,positionNumber_target
	'			,moveCount,maxStackCount
	' response : characterName
	'			,bagNumber_source,positionNumber_source
	'			,itemHash_source,stackedCount_source,itemSerial_source,instance_source,itemExpire_source
	'			,bagNumber_target,positionNumber_target
	'			,itemHash_target,stackedCount_target,itemSerial_target,instance_target,itemExpire_target
	'			,moveCount

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName _
		,bagNumber_source,positionNumber_source _
		,bagNumber_target,positionNumber_target _
		,moveCount,maxStackCount
	If UBound(parameters)<6 Then Call Error("not enough parameter")
	characterName			= Parameters(0)
	bagNumber_source		= Parameters(1)
	positionNumber_source	= Parameters(2)
	bagNumber_target		= Parameters(3)
	positionNumber_target	= Parameters(4)
	moveCount				= Parameters(5)
	maxStackCount			= Parameters(6)
	
	Dim itemHash_source,stackedCount_source,itemSerial_source,instance_source,itemExpire_source _
		,itemHash_target,stackedCount_target,itemSerial_target,instance_target,itemExpire_target

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.BankItems_Swap"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@bagNumber_source",adUnsignedTinyInt,adParamInput,,bagNumber_source)
		Call .AppendParam("@positionNumber_source",adUnsignedTinyInt,adParamInput,,positionNumber_source)
		Call .AppendParam("@bagNumber_target",adUnsignedTinyInt,adParamInput,,bagNumber_target)
		Call .AppendParam("@positionNumber_target",adUnsignedTinyInt,adParamInput,,positionNumber_target)		
		Call .AppendParam("@moveCount",adInteger,adParamInput,,moveCount)
		Call .AppendParam("@maxStackCount",adInteger,adParamInput,,maxStackCount)
		Call .AppendParam("@itemHash_source",adInteger,adParamOutput,,null)
		Call .AppendParam("@stackedCount_source",adInteger,adParamOutput,,null)
		Call .AppendParam("@itemSerial_source",adBigInt,adParamOutput,,null)
		Call .AppendParam("@instance_source",adBigInt,adParamOutput,,null)
		Call .AppendParam("@itemExpire_source",adDBTimeStamp,adParamOutput,,null)
		Call .AppendParam("@itemHash_target",adInteger,adParamOutput,,null)
		Call .AppendParam("@stackedCount_target",adInteger,adParamOutput,,null)
		Call .AppendParam("@itemSerial_target",adBigInt,adParamOutput,,null)
		Call .AppendParam("@instance_target",adBigInt,adParamOutput,,null)
		Call .AppendParam("@itemExpire_target",adDBTimeStamp,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If		
		itemHash_source		= sphn.GetParamValue("@itemHash_source")
		stackedCount_source	= sphn.GetParamValue("@stackedCount_source")
		itemSerial_source	= sphn.GetParamValue("@itemSerial_source")
		instance_source		= sphn.GetParamValue("@instance_source")
		itemExpire_source	= sphn.GetParamValue("@itemExpire_source")
		itemHash_target		= sphn.GetParamValue("@itemHash_target")
		stackedCount_target	= sphn.GetParamValue("@stackedCount_target")
		itemSerial_target	= sphn.GetParamValue("@itemSerial_target")
		instance_target		= sphn.GetParamValue("@instance_target")
		itemExpire_target	= sphn.GetParamValue("@itemExpire_target")
	End If
	set sphn = Nothing

	retString = characterName &_
			SEPARATOR & bagNumber_source & SEPARATOR & positionNumber_source &_
			SEPARATOR & itemHash_source & SEPARATOR & stackedCount_source &_
			SEPARATOR & itemSerial_source & SEPARATOR & instance_source & SEPARATOR & FormatDt(itemExpire_source,"SQL_TM") &_
			SEPARATOR & bagNumber_target & SEPARATOR & positionNumber_target &_
			SEPARATOR & itemHash_target & SEPARATOR & stackedCount_target &_
			SEPARATOR & itemSerial_target & SEPARATOR & instance_target & SEPARATOR & FormatDt(itemExpire_target,"SQL_TM") &_
			SEPARATOR & moveCount
			
	Call OK(retString)
%>