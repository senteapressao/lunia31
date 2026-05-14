<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : character name
	' response : nothing

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	
	' preparing request parameters	
	Dim channelName,characterName,connectionSerial
	channelName = REQUESTERID		
	characterName = Parameters(0)
	connectionSerial = Parameters(1)

	' sql query to retrieve Pvp
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_UpdatePvpConnection"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@isActive",adBoolean,adParamInput,,1)
		Call .AppendParam("@status",adVarWChar,adParamInput,80,null)
		Call .AppendParam("@stageEnd",adBoolean,adParamInput,,null)
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