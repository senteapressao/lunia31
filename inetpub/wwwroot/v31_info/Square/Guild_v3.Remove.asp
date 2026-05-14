<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : guildId,characterName
	' response : 0
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim guildId,characterName
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	guildId = Parameters(0)
	characterName = Parameters(1)
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Remove"
		Call .InitCommand()
		Call .AppendParam("@guildId",adInteger,adParamInput,,guildId)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	End If
	Set sph = Nothing

	Call Ok(0)
%>