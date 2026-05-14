<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request : guildId,guildPlayTime_add
	' response : guildId,guildPlayTime_old,guildPlayTime_new
	
	Dim i,j,k
	Dim sph,rs,retString,blsResult : blsResult = False

	' preparing request parameters
	Dim guildId,guildPlayTime_add
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	guildId = Parameters(0)
	guildPlayTime_add = Parameters(1)
	
	Dim guildPlayTime_old,guildPlayTime_new
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_AddGuildPlayTime"
		Call .InitCommand()
		Call .AppendParam("@guildId",adInteger,adParamInput,,guildId)
		Call .AppendParam("@guildPlayTime_add",adBigInt,adParamInput,,guildPlayTime_add)
		Call .AppendParam("@guildPlayTime_old",adBigInt,adParamOutput,,null)
		Call .AppendParam("@guildPlayTime_new",adBigInt,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		guildPlayTime_old = sph.GetParamValue("@guildPlayTime_old")
		guildPlayTime_new = sph.GetParamValue("@guildPlayTime_new")
	End If
	Set sph = Nothing
	
	retString = retString & guildId & SEPARATOR & guildPlayTime_old & SEPARATOR & guildPlayTime_new

	Call Ok(retString)
%>