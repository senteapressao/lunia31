<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request	: characterName
	' response	: lastMail,lastTrade
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim rs2
	
	' preparing request parameters
	Dim characterName
	If UBound(parameters)<0 Then Call Error("not enough parameter")	
	characterName	= Parameters(0)
	
	Dim lastMail,lastTrade
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.MailTradeVerifyLastSent"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@lastMail",adDBTimeStamp,adParamOutput,,null)
		Call .AppendParam("@lastTrade",adDBTimeStamp,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		lastMail = FormatDt(sphn.GetParamValue("@lastMail"),"SQL_TM")
		lastTrade = FormatDt(sphn.GetParamValue("@lastTrade"),"SQL_TM")
	End If
	set sphn = Nothing


	Call OK(characterName & SEPARATOR & lastMail & SEPARATOR & lastTrade)
%>