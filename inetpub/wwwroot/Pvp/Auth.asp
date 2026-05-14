<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : reservationSerial
	' response : characterName,virtualIdCode,classNumber,stageLevel,stageExp,pvpLevel,pvpExp,warLevel,warExp,instantStateFlags
	'			,rebirthCount,storedLevel,ladderPoint,dailyPlayCount,playCount,winCount
	'			,{stageGroupHash,accessLevel}
	'			,{itemHash,positionNumber,instance}
	'			,{bagNumber,positionNumber,itemHash,instance}

	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	Dim channelName,reservationSerial
	channelName = REQUESTERID
	reservationSerial = Clng(Parameters(0))

	Dim characterName,isBalanced
	' try to authorize
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_AuthPvpConnection"
		Call .InitCommand()
		Call .AppendParam("@reservationSerial",adBigInt,adParamInput,,reservationSerial)
		Call .AppendParam("@channelName",adVarWChar,adParamInput,50,channelName)
		Call .AppendParam("@characterName",adVarWChar,adParamOutput,50,null)
		Call .AppendParam("@isBalanced",adBoolean,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		characterName = sph.GetParamValue("@characterName")
		isBalanced = sph.GetParamValue("@isBalanced")
	End If
	set sph = Nothing
	retString = reservationSerial

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.InfoCharacterForPvp"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset()
	End with

	If blsResult Then
		Set rs = sphn.rs
	End If
	set sphn = Nothing

	If rs.Eof Then Call Error("no recordset")

	' basic : characterName,virtualIdCode,classNumber,stageLevel,stageExp,pvpLevel,pvpExp,instantStateFlags
	retString = retString & SEPARATOR & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) &_
			SEPARATOR & rs(3) & SEPARATOR & rs(4) & SEPARATOR & rs(5) &_
			SEPARATOR & rs(6) & SEPARATOR & rs(7) & SEPARATOR & rs(8) &_
			SEPARATOR & rs(9)
	
	' characterRebirth : rebirthCount,storedLevel
	retString = retString & SEPARATOR
	set rs = rs.NextRecordset
	If Not(rs.Eof) Then
		retString = retString & rs(0)  & SEPARATOR & rs(1)
	Else
		retString = retString & 0 & SEPARATOR & 0
	End If
	
	' ladder : ladderPoint,dailyPlayCount,playCount,winCount
	retString = retString & SEPARATOR
	set rs = rs.NextRecordset
	If Not(rs.Eof) Then
		'retString = retString & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & rs(3)
		retString = retString & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & rs(3)
	Else
		'retString = retString & 1500 & SEPARATOR & 0 & SEPARATOR & 0 & SEPARATOR & 0
		retString = retString & 1500 & SEPARATOR & 0 & SEPARATOR & 0 & SEPARATOR & 0
	End If

	' license : stageGroupHash,accessLevel
	i = 0
	retString = retString & SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then
			retString = retString & SUBSEPARATOR
		End If
		retString = retString & rs(0) & SUBSEPARATOR & rs(1)
		rs.MoveNext : i=i+1
	Loop

	' items : bagNumber,positionNumber,itemHash,instance,itemExpire
	i = 0
	retString = retString & SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then
			retString = retString & SUBSEPARATOR
		End If
		retString = retString & rs(0) & SUBSEPARATOR & rs(1) & SUBSEPARATOR & rs(2) & SUBSEPARATOR & rs(3) & SUBSEPARATOR & FormatDt(rs(4),"SQL_TM")
		rs.MoveNext : i=i+1
	Loop
	rs.Close

	' instantStateFlags
	' IsBalanced modify
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.UpdateInstantStateFlagsWithMask"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)		
		Call .AppendParam("@startBit",adUnsignedTinyInt,adParamInput,,8)
		Call .AppendParam("@bitLength",adUnsignedTinyInt,adParamInput,,1)
		If isBalanced Then
			Call .AppendParam("@value",adInteger,adParamInput,,1)
		Else
			Call .AppendParam("@value",adInteger,adParamInput,,0)
		End If
		blsResult = .ExecNoRecords()
	End with
	set sphn = Nothing

	Call Ok(retString)
%>