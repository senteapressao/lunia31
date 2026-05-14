<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()	
	' request  : tempId,characterName
	' response : tempId,characterName
	'			,classNumber,stageLevel,stageExp,pvpLevel,pvpExp,warLevel,warExp,gameMoney,bankMoney,skillPoint,instantStateFlags,createDate,characterSerial
	'			,addedSkillPoint
	'			,rebirthCount,storedLevel,storedSkillPoint,lastRebirthDate
	'			,ladderPoint,dailyPlayCount,playCount,winCount
	'			,achievementPoint
	'			,characterLicenses(subseparated: classNumber)
	'			,licenses(subseparated: stageGroupHash,accessLevel)
	'			,skilllicenses(subseparated: skillGroupHash)
	'			,Items(subseparated: itemHash,isOnBank,bagNumber,positionNumber,stackedCount,itemSerial,instance,expireDate)
	'			,skills(subseparated: skillGroupHash,skillLevel)
	'			,quickslots(subseparated: hash,position,isSkill,instance,expireDate)
	'			,bags(subseparated: bagNumber,isBank,expireDate)
	'			,bankBags(subseparated: bagNumber,expireDate)
	'			,pets(subseparated: petSerial,petName,petLevel,petExp,fullValue,isRare,rareProbability,fullRateSum,dtSum,petHash)
	'			,petItems(subseparated: petSerial,bagNumber,positionNumber,itemHash,stackedCount,itemSerial,instance,expireDate)
	'			,petTraining(subseparated: petSerial,itemHash,stackedCount,instance,expFactor,startTime,endTime)
	'			,emojiCommands(subseparated: command,value)
	'			,itemEffects(subseperated: license)
	'			,wingLicense(subseperated: wingType)
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim tempId,characterName
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	tempId	= Parameters(0)
	characterName	= Parameters(1)

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.InfoCharacterForStage"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset()
	End with
	
	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		Set rs = sphn.rs
	Else
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
	End If
	set sphn = Nothing

	Dim rs2
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.InfoCharacterForStage_Items"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset()
	End with
	
	If blsResult Then
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		Set rs2 = sphn.rs
	Else
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
	End If
	set sphn = Nothing
	
	If rs.Eof Then Call Error(1)	' not existing character

	Response.Write tempId & SEPARATOR & characterName
	
	' basic : classNumber,stageLevel,stageExp,pvpLevel,pvpExp,warLevel,warExp,gameMoney,bankMoney,skillPoint,instantStateFlags,createDate,characterSerial
	Response.Write SEPARATOR & rs(0) & SEPARATOR & rs(1) &_
			SEPARATOR & rs(2) & SEPARATOR & rs(3) & SEPARATOR & rs(4) &_
			SEPARATOR & rs(5) & SEPARATOR & rs(6) & SEPARATOR & rs(7) &_
			SEPARATOR & rs(8) & SEPARATOR & rs(9) & SEPARATOR & rs(10) &_
			SEPARATOR & FormatDt(rs(11),"SQL_TM") & SEPARATOR & rs(12)
	
	' addedSkillPoint : skillPoint
	set rs = rs.NextRecordset
	If Not(rs.Eof) Then
		Response.Write SEPARATOR & rs(0)
	Else
		Response.Write SEPARATOR & 0
	End If
	
	' characterRebirth : rebirthCount,storedLevel,skillPoint
	set rs = rs.NextRecordset
	If Not(rs.Eof) Then
		Response.Write SEPARATOR & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & FormatDt(rs(3),"SQL_TM") & SEPARATOR & FormatDt(rs(4),"SQL_TM")
	Else
		Response.Write SEPARATOR & 0 & SEPARATOR & 1 & SEPARATOR & 0 & SEPARATOR & FormatDt("1900-01-01","SQL_TM") & SEPARATOR & FormatDt("1900-01-01","SQL_TM")
	End If
	
	' ladder : ladderPoint,dailyPlayCount,playCount,winCount
	set rs = rs.NextRecordset
	If Not(rs.Eof) Then
		'Response.Write SEPARATOR & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & rs(3)
		Response.Write SEPARATOR & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & rs(3)
	Else
		Response.Write SEPARATOR & 3000 & SEPARATOR & 0 & SEPARATOR & 0 & SEPARATOR & 0
	End If

	' achievement : achievementScore
	set rs = rs.NextRecordset
	If Not(rs.Eof) Then
		Response.Write SEPARATOR & rs(0)
	Else
		Response.Write SEPARATOR & 0 
	End If
	
	' characterLicense : classNumber
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0)
		rs.MoveNext : i=i+1
	Loop

	' license : stageGroupHash,accessLevel
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0) & SUBSEPARATOR & rs(1)
		rs.MoveNext : i=i+1
	Loop
	
	' skilllicenses : skillGroupHash
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0)
		rs.MoveNext : i=i+1
	Loop

	' items : itemHash,isOnBank, bagNumber, positionNumber, stackedCount, itemSerial, instance, expireDate
	i = 0
	Response.Write SEPARATOR
	Do Until rs2.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs2(0) & SUBSEPARATOR & CInt(rs2(1)) & SUBSEPARATOR & rs2(2) &_
					SUBSEPARATOR & rs2(3) & SUBSEPARATOR & rs2(4) & SUBSEPARATOR & rs2(5) & SUBSEPARATOR & rs2(6) &_
					SUBSEPARATOR & FormatDt(rs2(7),"SQL_TM")
				
		rs2.MoveNext : i=i+1
	Loop
	rs2.Close

	' skills : skillGroupHash,skillLevel
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0) & SUBSEPARATOR & rs(1)
		rs.MoveNext : i=i+1
	Loop

	' quickslots : hash,position,isSkill,instance, itemExpire
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0) & SUBSEPARATOR & rs(1) &_
				SUBSEPARATOR & CInt(rs(2)) & SUBSEPARATOR & rs(3) & SUBSEPARATOR & FormatDt(rs(4),"SQL_TM")
		rs.MoveNext : i=i+1
	Loop

	' bags : bagNumber,isBank,expireDate
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0) & SUBSEPARATOR & CInt(rs(1))
		If IsNull(rs(2)) Then
			Response.Write SUBSEPARATOR & ""
		Else
			Response.Write SUBSEPARATOR & FormatDt(rs(2),"SQL_TM")
		End If
		rs.MoveNext : i=i+1
	Loop
	
	' bankBags : bagNumber,expireDate
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0)
		If IsNull(rs(1)) Then
			Response.Write SUBSEPARATOR & ""
		Else
			Response.Write SUBSEPARATOR & FormatDt(rs(1),"SQL_TM")
		End If
		rs.MoveNext : i=i+1
	Loop
	
	' pets : petSerial,petName,petLevel,petExp,fullValue,isRare,rareProbability,fullRateSum,dtSum,petHash,petEnchantserial
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0) & SUBSEPARATOR & rs(1) & SUBSEPARATOR & rs(2) &_
				SUBSEPARATOR & rs(3) & SUBSEPARATOR & rs(4) & SUBSEPARATOR & Int(rs(5)) &_
				SUBSEPARATOR & rs(6) & SUBSEPARATOR & rs(7) & SUBSEPARATOR & rs(8) &_
				SUBSEPARATOR & rs(9) & SUBSEPARATOR & rs(10)
		rs.MoveNext : i=i+1
	Loop
	
	' petItems : petSerial,bagNumber,positionNumber,itemHash,stackedCount,itemSerial,instance,expireDate
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0) & SUBSEPARATOR & rs(1) & SUBSEPARATOR & rs(2) &_
				SUBSEPARATOR & rs(3) & SUBSEPARATOR & rs(4) & SUBSEPARATOR & rs(5) &_
				SUBSEPARATOR & rs(6) & SUBSEPARATOR & FormatDt(rs(7),"SQL_TM")
		rs.MoveNext : i=i+1
	Loop
	
	' petTraining : petSerial,itemHash,stackedCount,instance,itemExpire,expFactor,startTime,endTime
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0) &_
				SUBSEPARATOR & rs(1) &_
				SUBSEPARATOR & rs(2) &_
				SUBSEPARATOR & rs(3) &_
				SUBSEPARATOR & FormatDt(rs(4),"SQL_TM") &_
				SUBSEPARATOR & rs(5) &_
				SUBSEPARATOR & FormatDt(rs(6),"SQL_TM") &_
				SUBSEPARATOR & FormatDt(rs(7),"SQL_TM")
		rs.MoveNext : i=i+1
	Loop	
	' emojiCommands : command,value
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0) &_
				SUBSEPARATOR & rs(1)
		rs.MoveNext : i=i+1
	Loop

	' itemEffects : license
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0)
		rs.MoveNext : i=i+1
	Loop

	' wingLicense : wingType
	i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0)
		rs.MoveNext : i=i+1
	Loop

	rs.Close	
%>