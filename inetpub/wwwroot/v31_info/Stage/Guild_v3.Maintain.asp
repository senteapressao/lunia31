<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : guildId,guildMemberId,guildPoint
	' response : guildLevel,guildPoint,shopStart,shopEnd
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim guildId,guildMemberId,guildPoint
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	guildId = Parameters(0)
	guildMemberId = Parameters(1)
	guildPoint = Parameters(2)
	
	Dim guildLevel,shopStart,shopEnd,guildPoint_o
	guildPoint_o = guildPoint
	
	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Maintain"
		Call .InitCommand()
		Call .AppendParam("@guildId",adInteger,adParamInput,,guildId)
		Call .AppendParam("@guildMemberId",adInteger,adParamInput,,guildMemberId)
		Call .AppendParam("@guildPoint",adInteger,adParamInputOutput,,guildPoint)
		Call .AppendParam("@guildLevel",adUnsignedTinyint,adParamOutput,,null)
		Call .AppendParam("@shopStart",adDBTimeStamp,adParamOutput,,null)
		Call .AppendParam("@shopEnd",adDBTimeStamp,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		guildPoint = sph.GetParamValue("@guildPoint")
		guildLevel = sph.GetParamValue("@guildLevel")
		shopStart = sph.GetParamValue("@shopStart")
		shopEnd = sph.GetParamValue("@shopEnd")
	End If
	Set sph = Nothing

	Call Ok(guildLevel & SEPARATOR & guildPoint & SEPARATOR & FormatDt(shopStart,"SQL_TM") & SEPARATOR & FormatDt(shopEnd,"SQL_TM") & SEPARATOR & guildPoint_o)
%>