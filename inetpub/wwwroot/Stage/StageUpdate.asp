<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : roomNumber,password,stageGroupHash,accessLevel,step,players,isLocked,capacity
	' response : nothing

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim serverName,roomNumber,password,stageGroupHash,accessLevel,step,players,isLocked,capacity
	If UBound(parameters)<6 Then Call Error("not enough parameter")
	serverName		= REQUESTERID
	roomNumber		= Parameters(0)
	password		= Parameters(1)
	stageGroupHash	= Parameters(2)
	accessLevel		= Parameters(3)
	step			= Parameters(4)
	players			= Parameters(5)
	isLocked		= Parameters(6)
	capacity		= Parameters(7)
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_UpdateStage"
		Call .InitCommand()
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
		Call .AppendParam("@roomNumber",adInteger,adParamInput,,roomNumber)
		Call .AppendParam("@stageGroupHash",adInteger,adParamInput,,stageGroupHash)
		Call .AppendParam("@accessLevel",adSmallInt,adParamInput,,accessLevel)
		Call .AppendParam("@step",adTinyInt,adParamInput,,step)
		Call .AppendParam("@password",adVarWChar,adParamInput,80,password)
		Call .AppendParam("@isLocked",adBoolean,adParamInput,,isLocked)
		Call .AppendParam("@capacity",adSmallInt,adParamInput,,capacity)
		Call .AppendParam("@players",adSmallInt,adParamInput,,players)
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

	Call Ok("Ok")
%>