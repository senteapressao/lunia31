<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : temporary key, password, charactername, teamNumber
	' response : temporary key, stage server address, port, key code

	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	Dim tempKey,stageName,characterName,teamNumber,instantFlagForPvp
	Dim channelName,serverName,roomNumber,address,port,reservationSerial,isSpectator

	channelName = REQUESTERID
	If UBound(parameters)<3 Then Call Error("not enough parameter")
	tempKey = parameters(0)
	stageName = parameters(1)
	characterName = parameters(2)
	teamNumber = parameters(3)
	instantFlagForPvp = parameters(4)

	instantFlagForPvp = instantFlagForPvp AND Int("&H00FC00")
	instantFlagForPvp = instantFlagForPvp / 2^9

	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_GetActiveStageByPassword"
		Call .InitCommand()
		Call .AppendParam("@password",adVarWChar,adParamInput,80,stageName)
		Call .AppendParam("@serverName",adVarWChar,adParamOutput,50,null)
		Call .AppendParam("@roomNumber",adInteger,adParamOutput,,null)
		Call .AppendParam("@stageGroupHash",adInteger,adParamOutput,,null)
		Call .AppendParam("@accessLevel",adSmallInt,adParamOutput,,null)
		Call .AppendParam("@isLocked",adBoolean,adParamOutput,,null)
		Call .AppendParam("@address",adVarChar,adParamOutput,50,null)
		Call .AppendParam("@port",adInteger,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		response.write sph.frk_strErrorText
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		serverName = sph.GetParamValue("@serverName")
		roomNumber = sph.GetParamValue("@roomNumber")
		address = sph.GetParamValue("@address")
		port = sph.GetParamValue("@port")
	End If
	set sph = Nothing

	retString = stageName & SEPARATOR & address & SEPARATOR & port
	
	If teamNumber<255 Then	' Player
		isSpectator = 0
	Else	' Spectator
		isSpectator = 1
	End If

	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_JoinStageForPvp"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
		Call .AppendParam("@roomNumber",adInteger,adParamInput,,roomNumber)
		Call .AppendParam("@password",adVarWChar,adParamInput,80,stageName)
		Call .AppendParam("@teamNumber",adInteger,adParamInput,,teamNumber)
		Call .AppendParam("@reservationSerial",adBigInt,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
	'	if sph.frk_n4ErrorCode>0 Then
	'		Call Error(sph.frk_n4ErrorCode)
	'	Else
	'		Call Error(sph.frk_strErrorText)
	'	End If
		Call Error(10)
	Else
		reservationSerial = sph.GetParamValue("@reservationSerial")
	End If
	set sph = Nothing

	retString = retString & SEPARATOR & reservationSerial
	
	' instantStateFlags
	' IsSpectator modify
	' IsPetNotUsable,IsEquipSetRewardDisable,IsRebirthSkillDisable,IsInvincibleAfterRevive,StatusLimit modify
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.UpdateInstantStateFlagsWithMask"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)		
		Call .AppendParam("@startBit",adUnsignedTinyInt,adParamInput,,9)
		Call .AppendParam("@bitLength",adUnsignedTinyInt,adParamInput,,7)
		Call .AppendParam("@value",adInteger,adParamInput,,(instantFlagForPvp OR isSpectator))
		blsResult = .ExecNoRecords()
	End with
	set sphn = Nothing

	Call Ok(retString)	
%>