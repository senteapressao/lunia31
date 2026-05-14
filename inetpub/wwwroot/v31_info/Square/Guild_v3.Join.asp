<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : guildId,characterName,grade
	' response : guildId
	
	Dim i,j,k
	Dim sph,rs,rs2,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim guildId,characterName,grade
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	guildId = Parameters(0)
	characterName = Parameters(1)
	grade = Parameters(2)
	
	Dim guildName	
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Join"
		Call .InitCommand()
		Call .AppendParam("@guildId",adInteger,adParamInput,,guildId)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@grade",adUnsignedTinyint,adParamInput,,grade)
		Call .AppendParam("@guildName",adVarWChar,adParamOutput,14,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		guildName = sph.GetParamValue("@guildName")
	End If
	Set sph = Nothing

	Call Ok(guildId & SEPARATOR & guildName)
%>