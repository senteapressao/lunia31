<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request	: characterSerial,statisticsKey,statisticsValue 2Oracle0#8#61865383
	' response	: 
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	
	Dim rowData
	rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)
	
	
	' preparing request parameters
	Dim characterSerial,statisticsKey,statisticsValue
	If UBound(parameters)<2 Then Call Error("not enough parametergg")	
	characterSerial	= Parameters(0)
	statisticsKey	= Parameters(1)
	statisticsValue	= Parameters(2)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.statisticsSet"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterSerial",adBigInt,adParamInput,,characterSerial)
		Call .AppendParam("@statisticsKey",adVarChar,adParamInput,50,statisticsKey)
		Call .AppendParam("@statisticsValue",adVarChar,adParamInput,4000,statisticsValue)
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

	Call Ok(0)
%>