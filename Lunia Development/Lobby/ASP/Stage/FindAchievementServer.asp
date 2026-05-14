<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	' response : characterName,serverName,serverIP,clientIP,serverPort,clientPort

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	Dim characterName
	Dim serverName,serverIP,clientIP,serverPort,clientPort
	characterName	= Parameters(0)

	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_FindAchievementConnection"
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
			if sph.frk_n4ErrorCode>0 Then
				Call Error(sph.frk_n4ErrorCode)
			Else
				Call Error(sph.frk_strErrorText)
			End If
		End If
	End If
	serverName	= sph.GetParamValue("@serverName")
	serverIP	= sph.GetParamValue("@serverIP")
	clientIP	= sph.GetParamValue("@clientIP")
	serverPort	= sph.GetParamValue("@serverPort")
	clientPort	= sph.GetParamValue("@clientPort")

	retString = characterName & SEPARATOR & serverName & SEPARATOR & serverIP & SEPARATOR & clientIP & SEPARATOR & serverPort & SEPARATOR & clientPort
	set sph = Nothing

	Call Ok(retString)
%>
