<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request	: accountName,characterName
	' response	: accountName,characterName,characterSerial,isFamilyJoined
	'			serverName,serverIP,clientIP,serverPort,clientPort

	Dim i,j,k
	Dim sph,sphn,rs,retString,blsResult : blsResult = False

	' preparing request parameters
	Dim accountName,characterName
	Dim serverName,serverIP,clientIP,serverPort,clientPort
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	accountName		= Parameters(0)
	characterName	= Parameters(1)
	serverName		= REQUESTERID'"TESTETELES"
	
	Dim characterSerial
	Dim isFamilyJoined : isFamilyJoined = 0

	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_TerminateAchievementConnection"
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

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.SelectCharacter"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
		Call .AppendParam("@characterSerial",adBigInt,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		characterSerial = sphn.GetParamValue("@characterSerial")
	End If
	set sphn = Nothing
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_FindAchievementServer"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@serverName",adVarWChar,adParamOutput,50,null)
		Call .AppendParam("@serverIP",adVarChar,adParamOutput,50,null)
		Call .AppendParam("@clientIP",adVarChar,adParamOutput,50,null)
		Call .AppendParam("@serverPort",adInteger,adParamOutput,,null)
		Call .AppendParam("@clientPort",adInteger,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		Dim temp_sph_frk_n4ErrorCode,temp_sph_frk_strErrorText
		temp_sph_frk_n4ErrorCode	= sph.frk_n4ErrorCode
		temp_sph_frk_strErrorText	= sph.frk_strErrorText

		set sph = Nothing

		Set sphn = new SPHelper_NoTran
		with sphn
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = characterDBconnectionString
			.SPName = "dbo.DeselectCharacter"
			Call .InitCommand()
			Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
			Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
			blsResult = .ExecNoRecords()
		End with

		set sphn = Nothing

		if temp_sph_frk_n4ErrorCode>0 Then
			Call Error(temp_sph_frk_n4ErrorCode)
		Else
			Call Error(temp_sph_frk_strErrorText)
		End If
	Else
		serverName			= sph.GetParamValue("@serverName")
		serverIP			= sph.GetParamValue("@serverIP")
		clientIP			= sph.GetParamValue("@clientIP")
		serverPort			= sph.GetParamValue("@serverPort")
		clientPort			= sph.GetParamValue("@clientPort")
	End If
	set sph = Nothing
	
	' family Online
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Family_Online"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@isOnline",adBoolean,adParamInput,,1)
		Call .AppendParam("@isFamilyJoined",adBoolean,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with
	
	If blsResult=False Then
	Else
		isFamilyJoined = sph.GetParamValue("@isFamilyJoined")
		isFamilyJoined = Abs(CInt(isFamilyJoined))
	End If
	Set sph = Nothing
	
	retString = accountName & SEPARATOR & characterName & SEPARATOR & characterSerial & SEPARATOR & isFamilyJoined
	
	' Event_PlayTime
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.Event_InitPlayTime"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecNoRecords()
	End with
	Set sphn = nothing
	
	retString = retString & SEPARATOR & serverName & SEPARATOR & serverIP & SEPARATOR & clientIP & SEPARATOR & serverPort & SEPARATOR & clientPort

	Call Ok(retString)
%>