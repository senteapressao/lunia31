<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : accountName,characterName[,why]
	' response : accountName

	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim serverName,accountName,characterName,why	
	serverName = REQUESTERID
	accountName = Parameters(0)
	characterName = Parameters(1)
	If UBound(parameters)>1 Then
		why = Parameters(2)
	Else
		why = ""
	End If
	
	' call Stored procedure	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.TerminateAccount"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
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

	' call Stored procedure
	'Set sph = new SPHelper
	'with sph
		'.DEBUG = False
		'Set .cmd = Command
		'.ConnStr = stageDBconnectionString
		'.SPName = "dbo.public_TerminateConnection"
		'Call .InitCommand()
		'Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		'blsResult = .ExecNoRecords()
	'End with
	'set sph = Nothing

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
	set sph = Nothing
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_UpdateLatestLogout"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecNoRecords()
	End with
	set sph = Nothing
	
	' family Offline
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Family_Online"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@isOnline",adBoolean,adParamInput,,0)
		Call .AppendParam("@isFamilyJoined",adBoolean,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with
	Set sph = Nothing
	
	Call Ok(accountName)
%>
