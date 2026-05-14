<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request	: characterSerial,statisticsKey,moduleKey
	' response	: statisticsKey,statisticsValue,moduleKey
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterSerial,statisticsKey,moduleKey
	If UBound(parameters)<2 Then Call Error("not enough parameter")	
	characterSerial	= Parameters(0)
	statisticsKey	= Parameters(1)
	moduleKey		= Parameters(2)
	
	Dim statisticsValue

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.statisticsGet"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterSerial",adBigInt,adParamInput,,characterSerial)
		Call .AppendParam("@statisticsKey",adVarChar,adParamInput,50,statisticsKey)
		Call .AppendParam("@statisticsValue",adVarChar,adParamOutput,4000,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		statisticsValue = sphn.GetParamValue("@statisticsValue")
	End If
	set sphn = Nothing

	retString = statisticsKey & SEPARATOR & statisticsValue & SEPARATOR & moduleKey

	Call OK(retString)
%>