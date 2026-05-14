<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : accountName,characterName
	' response : accountName,characterName,money

	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim serverName,accountName,characterName,money
	serverName = REQUESTERID
	accountName = Parameters(0)
	characterName = Parameters(1)
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_CheckJoined"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecNoRecords()
	End with
	
	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(3)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	End If
	Set sph = Nothing
	
	' delete familyMember
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Family_DeleteMember"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecNoRecords()
	End with
	Set sph = Nothing
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.DeleteCharacter"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@money",adInteger,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with
	Set sph = Nothing

	If blsResult Then
		money = sphn.GetParamValue("@money")
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
	End If
	set sphn = Nothing	
	
	Call Ok(accountName & SEPARATOR & characterName & SEPARATOR & money)
%>
