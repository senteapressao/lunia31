<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : character name, isWin, pvpSubInfo
	' response : character name, isWin, pvpSubInfo

	Dim characterName,isWin,pvpSubInfo
	
	characterName=Parameters(0)
	isWin=Parameters(1)
	pvpSubInfo=Parameters(2)

	Dim i,j,k
	Dim sph,RS,blsResult : blsResult = False
	Dim retString : retString = ""	

	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_UpdatePvpConnection"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@isActive",adBoolean,adParamInput,,null)
		Call .AppendParam("@status",adVarWChar,adParamInput,80,null)
		Call .AppendParam("@stageEnd",adBoolean,adParamInput,,1)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	End If
	set sph = Nothing

	retString = characterName & SEPARATOR & isWin & SEPARATOR & pvpSubInfo

	Call Ok(retString)
%>