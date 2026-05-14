<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	' response : characterName,virtualIdCode,classNumber,stageLevel,stageExp,pvpLevel,pvpExp,warLevel,warExp,instantStateFlags
	'			,rebirthCount,storedLevel,ladderPoint,dailyPlayCount,playCount,winCount
	'			,{stageGroupHash,accessLevel}
	'			,{bagNumber,positionNumber,itemHash,instance}
	'			,status

	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	If UBound(parameters)<0 Then Call Error("not enough parameter")

	Dim channelName,characterName,status
	channelName = REQUESTERID
	characterName = Parameters(0)
	
	
	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_FinishPvpStage"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@channelName",adVarWChar,adParamInput,50,channelName)
		Call .AppendParam("@status",adVarWChar,adParamOutput,80,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		status = sph.GetParamValue("@status")
	End If
	set sph = Nothing

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

	' basic : characterName,virtualIdCode,classNumber,stageLevel,stageExp,pvpLevel,pvpExp,warLevel,warExp,instantStateFlags
	retString = rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) &_
			SEPARATOR & rs(3) & SEPARATOR & rs(4) & SEPARATOR & rs(5) &_
			SEPARATOR & rs(6) & SEPARATOR & rs(7) & SEPARATOR & rs(8) &_ 
			SEPARATOR & rs(9)
			
	' rebirth infomation
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
		retString = retString & 1500 & SEPARATOR & 10 & SEPARATOR & 0 & SEPARATOR & 0
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

	' items : bagNumber,positionNumber,itemHash,instance
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

	retString = retString & SEPARATOR & status

	Call Ok(retString)
%>
