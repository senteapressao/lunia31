<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : temporary key, password, charactername, teamNumber
	' response : temporary key, stage server address, port, key code

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	Dim stageName
	Dim stageGroupHash,accessLevel,isLocked

	stageName = parameters(0)

	If stageName = "" Then Call Error(1) ' invalid password

	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_GetActiveStageByPassword"
		Call .InitCommand()
		Call .AppendParam("@password",adVarWChar,adParamInput,80,stageName)
		Call .AppendParam("@serverName",adVarWChar,adParamOutput,50,null)
		Call .AppendParam("@roomNumber",adInteger,adParamOutput,,null)
		Call .AppendParam("@stageGroupHash",adInteger,adParamOutput,,null)
		Call .AppendParam("@accessLevel",adSmallInt,adParamOutput,,null)
		Call .AppendParam("@isLocked",adBoolean,adParamOutput,,null)
		Call .AppendParam("@address",adVarChar,adParamOutput,50,null)
		Call .AppendParam("@port",adInteger,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			'Call Error(sph.frk_n4ErrorCode)
			Call Error(2)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		stageGroupHash = sph.GetParamValue("@stageGroupHash")
		accessLevel = sph.GetParamValue("@accessLevel")
		isLocked = sph.GetParamValue("@isLocked")
	End If
	set sph = Nothing

	If isLocked Then
		Call Error(3)
	End If

	retString = stageGroupHash & SEPARATOR & accessLevel
	
	Call Ok(retString)
%>