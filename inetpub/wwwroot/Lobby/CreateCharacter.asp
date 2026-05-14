<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : accountName,characterName,classNumber,validateCharacterLicense
	' response : characterName,CharacterSerial,virtualIdCode,classNumber,stageLevel,stageExp,pvpLevel,pvpExp,warLevel,warExp,instantStateFlags
	'			,rebirthCount,storedLevel
	'			,{stageGroupHash,accessLevel}
	'			,{bagNumber,positionNumber,itemHash,instance,itemExpire}
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim accountName,characterName,classNumber,validateCharacterLicense
	If UBound(parameters)<3 Then Call Error("not enough parameter")	
	accountName					= Parameters(0)
	characterName				= Parameters(1)
	classNumber					= Parameters(2)
	validateCharacterLicense	= Parameters(3)

	If Instr(characterName,"?")>0 Then Call Error(3)

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.CreateCharacter"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@classNumber",adSmallInt,adParamInput,,classNumber)
		Call .AppendParam("@validateCharacterLicense",adBoolean,adParamInput,,validateCharacterLicense)
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
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.InfoCharacterForLobby"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset()
	End with	
	
	If blsResult Then
		Set rs = sphn.rs
	End If
	set sphn = Nothing

	If rs.Eof Then Call Error("no recordset")
	
	' basic : characterName,CharacterSerial,virtualIdCode,classNumber,stageLevel,stageExp,pvpLevel,pvpExp,warLevel,warExp,instantStateFlags
	retString = rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) &_
			SEPARATOR & rs(3) & SEPARATOR & rs(4) & SEPARATOR & rs(5) &_
			SEPARATOR & rs(6) & SEPARATOR & rs(7) & SEPARATOR & rs(8) &_
			SEPARATOR & rs(9) & SEPARATOR & rs(10)
				
	' rebirth infomation
	retString = retString & SEPARATOR
	set rs = rs.NextRecordset
	If Not(rs.Eof) Then
		retString = retString & rs(0)  & SEPARATOR & rs(1)
	Else
		retString = retString & 0 & SEPARATOR & 0
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
				
	Call Ok(retString)
%>
