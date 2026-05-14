<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : character name
	' response : character name

	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False

	Dim characterName	
	characterName=Parameters(0)

	' instantStateFlags
	' IsBalanced,IsSpectator off
	' IsPetNotUsable,IsEquipSetRewardDisable,IsRebirthSkillDisable,IsInvincibleAfterRevive off,StatusLimit=0
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.UpdateInstantStateFlagsWithMask"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)		
		Call .AppendParam("@startBit",adUnsignedTinyInt,adParamInput,,8)
		Call .AppendParam("@bitLength",adUnsignedTinyInt,adParamInput,,8)
		Call .AppendParam("@value",adInteger,adParamInput,,0)
		blsResult = .ExecNoRecords()
	End with
	set sphn = Nothing
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_TerminatePvpConnection"
		Call .InitCommand()
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
	set sph = Nothing

	Call Ok(characterName)
%>