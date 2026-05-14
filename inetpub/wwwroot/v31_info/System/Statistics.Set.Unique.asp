<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request	: characterSerial(=0),statisticsKey,statisticsValue
	' response	: isSuccess(1:Success/0:Fail),statisticsKey
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterSerial,statisticsKey,statisticsValue
	If UBound(parameters)<2 Then Call Error("not enough parameter")	
	characterSerial	= Parameters(0)
	statisticsKey	= Parameters(1)
	statisticsValue	= Parameters(2)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.statisticsSetUnique"
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
		if ret=1 Then
			Call OK(0 & SEPARATOR & statisticsKey)
		Elseif ret<>0 Then
			Call Error(ret)
		End If
	End If
	set sphn = Nothing

	Call OK(1 & SEPARATOR & statisticsKey)
%>