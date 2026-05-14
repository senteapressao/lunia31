<!--#include file="../include/connect.asp"-->
<!--#include file="../include/common.asp"-->
<!--#include file="../include/helpers.asp"-->
<%
	Dim serverCode,accountName,characterName,classNumber
	Dim str,dex,vit,int, maxHP, maxMP, dmgMax, dmgMin
	Dim dmgSkillFactor, magicCritMultiplyFactor, magicCritProbability
	Dim physCritMultiplyFactor, physCritProbability,skillCooldown
	Dim deathCount,fatal,dmgSum,healSum,hitdmgsum,mpspendsum,dpsNo
	Dim stackedPartyBattleTime,stageGroupHash,accessLevel,difficulty

	
    If Ubound(Parameters)<26 Then Call Error("not enough parameter")
	
	serverCode = Parameters(0)
	accountName = Parameters(1)						'Teles
	characterName = Parameters(2)                   'Teles
	classNumber = Parameters(3)                     '6
	str = Parameters(4)                             '24408
	dex = Parameters(5)                             '8390
	vit = Parameters(6)                             '10114
	int = Parameters(7)                             '8013
	maxHP = Parameters(8)                           '91941
	maxMP = Parameters(9)                           '71524
	dmgMax = Parameters(10)                         '9184
	dmgMin = Parameters(11)                         '8136
	dmgSkillFactor = Parameters(12)                 '57.76 f
	magicCritMultiplyFactor = Parameters(13)        '2.70 f
	magicCritProbability = Parameters(14)           '.27 f
	physCritMultiplyFactor = Parameters(15)         '12.11 f
	physCritProbability = Parameters(16)   		    '.39 f
	skillCooldown = Parameters(17)                  '.23 f
	deathCount = Parameters(18)                     '0
	fatal = Parameters(19)                          '0
	dmgSum = Parameters(20)                         '4
	healSum = Parameters(21)                        '0
	hitdmgsum = Parameters(22)                      '292400
	mpspendsum = Parameters(23)                     '0
	dpsNo = Parameters(24)                          '14700
	stackedPartyBattleTime = Parameters(25)         '19.89 f
	stageGroupHash = Parameters(26)                 '474497
	accessLevel = Parameters(27)                    '0
	difficulty = Parameters(28)						'1
	
	' call Stored procedure
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = false
		Set .cmd = Command
		.ConnStr = ConnStrCharacter
		.SPName = "dbo.Rank_Dps"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@classNumber",adInteger,adParamInput,,classNumber)
		Call .AppendParam("@str",adInteger,adParamInput,,str)
		Call .AppendParam("@dex",adInteger,adParamInput,,dex)
		Call .AppendParam("@vit",adInteger,adParamInput,,vit)
		Call .AppendParam("@int",adInteger,adParamInput,,int)
		Call .AppendParam("@maxHP",adInteger,adParamInput,,maxHP)
		Call .AppendParam("@maxMP",adInteger,adParamInput,,maxMP)
		Call .AppendParam("@dmgMax",adDouble,adParamInput,,dmgMax)
		Call .AppendParam("@dmgMin",adDouble,adParamInput,,dmgMin)
		Call .AppendParam("@dmgSkillFactor",adDouble,adParamInput,,dmgSkillFactor)
		Call .AppendParam("@magicCritMultiplyFactor",adDouble,adParamInput,,magicCritMultiplyFactor)
		Call .AppendParam("@magicCritProbability",adDouble,adParamInput,,magicCritProbability)
		Call .AppendParam("@physCritMultiplyFactor",adDouble,adParamInput,,physCritMultiplyFactor)
		Call .AppendParam("@physCritProbability",adDouble,adParamInput,,physCritProbability)
		Call .AppendParam("@skillCooldown",adDouble,adParamInput,,skillCooldown)
		
		Call .AppendParam("@deathCount",adInteger,adParamInput,,deathCount)
		Call .AppendParam("@fatal",adInteger,adParamInput,,fatal)
		Call .AppendParam("@dmgSum",adInteger,adParamInput,,dmgSum)
		Call .AppendParam("@healSum",adInteger,adParamInput,,healSum)
		Call .AppendParam("@hitdmgsum",adInteger,adParamInput,,hitdmgsum)
		Call .AppendParam("@mpspendsum",adInteger,adParamInput,,mpspendsum)
		
		Call .AppendParam("@dpsNo",adInteger,adParamInput,,dpsNo)
		Call .AppendParam("@stackedPartyBattleTime",adDouble,adParamInput,,stackedPartyBattleTime)
		
		Call .AppendParam("@stageGroupHash",adInteger,adParamInput,,stageGroupHash)
		Call .AppendParam("@accessLevel",adInteger,adParamInput,,accessLevel)
		Call .AppendParam("@difficulty",adInteger,adParamInput,,difficulty)
		blsResult = .ExecRecordset()
	End with

	ret = sphn.GetParamValue("RETURN_VALUE")
	if ret<>0 Then
		Call Error(ret)
	End If
	set sphn = Nothing
	
	Call Ok("Alright bb, do ya want some juicy juice?")
%>