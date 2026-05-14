<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : accountName,characterName,instantStateFlags
	' response : 
		
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False

	Dim accountName,characterName,instantStateFlags
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	accountName = Parameters(0)
	characterName = Parameters(1)
	instantStateFlags = Parameters(2)
	
	InstantStateFlags = InstantStateFlags AND Int("&H0000FF")
	
	' instantStateFlags
	' IsPcRoom,IsTestCharacter,IsInvalid,IsAdmin,IsGuestID modify
	' IsBalanced,IsSpectator off
	' IsPetNotUsable,IsEquipSetRewardDisable,IsRebirthSkillDisable,IsInvincibleAfterRevive off,StatusLimit=0
	' EquipmentSet,CashEquipmentSet not modify
	' LastSquareLocation not modify
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.UpdateInstantStateFlagsWithMask"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@startBit",adUnsignedTinyInt,adParamInput,,0)
		Call .AppendParam("@bitLength",adUnsignedTinyInt,adParamInput,,16)
		Call .AppendParam("@value",adInteger,adParamInput,,InstantStateFlags)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
	End If
	set sphn = Nothing

	Call Ok("Ok")
%>