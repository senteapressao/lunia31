<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request : guildId,guildMemberId,exp,guildLevel,nextLevelExp
	' response : guildLevel,guildExp,requestGuildExp
	
	Dim i,j,k
	Dim sph,rs,retString,blsResult : blsResult = False

	' preparing request parameters
	Dim guildId,guildMemberId,guildExp
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	guildId = Parameters(0)
	guildMemberId = Parameters(1)
	guildExp = Parameters(2)
	
	Dim guildLevel,guildExp_o,isLevelUp
	guildExp_o = guildExp
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_ExpUp"
		Call .InitCommand()
		Call .AppendParam("@guildId",adInteger,adParamInput,,guildId)
		Call .AppendParam("@guildMemberId",adInteger,adParamInput,,guildMemberId)
		Call .AppendParam("@guildExp",adInteger,adParamInputOutput,,guildExp)
		Call .AppendParam("@guildLevel",adUnsignedTinyint,adParamOutput,,null)
		Call .AppendParam("@isLevelUp",adBoolean,adParamOutput,,null)
		blsResult = .ExecRecordset()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		guildExp = sph.GetParamValue("@guildExp")
		guildLevel = sph.GetParamValue("@guildLevel")
		isLevelUp = sph.GetParamValue("@isLevelUp")
	End If
	Set sph = Nothing
	
	retString = retString & guildLevel & SEPARATOR & guildExp & SEPARATOR & guildExp_o
	If isLevelUp Then
		retString = retString & SEPARATOR & 1
	Else
		retString = retString & SEPARATOR & 0
	End If

	Call Ok(retString)
%>