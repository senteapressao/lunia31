<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	'parameters = Split("683069284994228101210-1.00000039936TelesCrazy",SEPARATOR)
	' request  : temporary key, stage group hash, access level of the stage group, state flags(see remarks), list of member characters
	' response : temporary key, stage server address, port, list of key codes
	'
	' member characters must be ordered 
	' flags is a integer that contains number of teams, private match flags, ...
	'
	' from lowest
	' 0~4 bits; number of teams
	' 5	  bit ; is private
	'

	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	Dim tempKey,stageGroupHash,accessLevel,step,stageStates,teamCapacity,battleGroundKillCount,battleGroundLimitTime,instantFlagForPvp
	ReDim characters(Ubound(parameters)-9)
	Dim channleName,stageName,teamCount,playerCount,teamNumber,isSpectator
	Dim serverName,roomNumber,address,port,reservationSerial
	channleName = REQUESTERID
	If UBound(parameters)<6 Then Call Error("not enough parameter")
	tempKey = parameters(0)
	stageGroupHash = parameters(1)
	accessLevel = parameters(2)
	step = parameters(3)
	stageStates = parameters(4)
	teamCapacity = parameters(5)
	battleGroundKillCount = parameters(6)
	battleGroundLimitTime = CDbl(parameters(7)) 'forced cdbl here
	instantFlagForPvp = parameters(8)
	j=0
	For i=9 To Ubound(parameters)
		If parameters(i)<>"" Then
			characters(j) = parameters(i)
			j=j+1
		End If
	Next
	' TODO: parameter validation
	stageName = "@"& FormatDt(Now(),"ISO_TM") &"@"& channleName &"@"& tempKey
	teamCount = GetNumberOfTeam(stageStates)
	playerCount = teamCount * teamCapacity

	instantFlagForPvp = instantFlagForPvp AND Int("&H00FC00")
	instantFlagForPvp = instantFlagForPvp / 2^9


	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = false
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_ReserveStageForPvp"
		Call .InitCommand()
		Call .AppendParam("@stageGroupHash",adInteger,adParamInput,,stageGroupHash)
		Call .AppendParam("@accessLevel",adSmallInt,adParamInput,,accessLevel)
		Call .AppendParam("@step",adTinyInt,adParamInput,,step)
		Call .AppendParam("@password",adVarWChar,adParamInput,80,stageName)
		Call .AppendParam("@capacity",adInteger,adParamInput,,playerCount)
		Call .AppendParam("@stageStates",adInteger,adParamInput,,stageStates)
		Call .AppendParam("@battleGroundKillCount",adInteger,adParamInput,,battleGroundKillCount)
		Call .AppendParam("@battleGroundLimitTime",adSingle,adParamInput,,battleGroundLimitTime)
		Call .AppendParam("@createSerial",adBigInt,adParamInput,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characters(0)) 'hotfix for lead maybe?
		Call .AppendParam("@serverName",adVarWChar,adParamOutput,50,null)
		Call .AppendParam("@roomNumber",adInteger,adParamOutput,,null)
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

	For i=0 To Ubound(characters)
		teamNumber = Int(i/teamCapacity)+1

		Set sph = new SPHelper
		with sph
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = stageDBconnectionString
			.SPName = "dbo.public_JoinStageForPvp"
			Call .InitCommand()
			Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characters(i))
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
		' IsPetNotUsable,IsEquipSetRewardDisable,IsRebirthSkillDisable,IsInvincibleAfterRevive,StatusLimit modify
		Set sphn = new SPHelper_NoTran
		with sphn
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = characterDBconnectionString
			.SPName = "dbo.UpdateInstantStateFlagsWithMask"
			Call .InitCommand()
			Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characters(i))
			Call .AppendParam("@startBit",adUnsignedTinyInt,adParamInput,,9)
			Call .AppendParam("@bitLength",adUnsignedTinyInt,adParamInput,,7)
			Call .AppendParam("@value",adInteger,adParamInput,,instantFlagForPvp)
			blsResult = .ExecNoRecords()
		End with
		set sphn = Nothing
	Next

	Call Ok(retString)
%>
<%
	Function SHL(Value, n) 
		Dim Tmp 
		Dim i 
		Tmp = Value 
		For i = 0 To n - 1 
			Tmp = Tmp * 2 
		Next 
		SHL = CLng(Tmp) 
	End Function 

	Function SHR(Value, n) 
		Dim Tmp 
		Dim i 
		Tmp = Value 
		For i = 0 To n - 1 
			Tmp = Tmp \ 2 
		Next 
		SHR = CLng(Tmp) 
	End Function 
	
	Function GetNumberOfTeam(value)
		const BitMast = &H0F
		GetNumberOfTeam = value AND BitMast
	End Function

	Function GetIsPrivate(value)
		const BitMast = &H10
		GetIsPrivate = SHR(value AND BitMast, 4)
	End Function
%>
