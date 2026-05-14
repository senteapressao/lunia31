<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : marketId
	' response : marketId,characterName,avgPrice,count

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim marketId
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	marketId = Parameters(0)

	Dim characterName,avgPrice,count

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.Market_Info"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@marketId",adBigInt,adParamInput,,marketId)
		Call .AppendParam("@characterName",adVarWChar,adParamOutput,50,null)
		Call .AppendParam("@avgPrice",adBigInt,adParamOutput,,null)
		Call .AppendParam("@count",adInteger,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		characterName = sphn.GetParamValue("@characterName")
		avgPrice = sphn.GetParamValue("@avgPrice")
		count = sphn.GetParamValue("@count")
	End If
	set sphn = Nothing

	retString = marketId & SEPARATOR & characterName &_
			SEPARATOR & avgPrice & SEPARATOR & count

	Call OK(retString)
%>