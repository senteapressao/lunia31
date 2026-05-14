<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init() 
	' request  : serverName,serverIP,clientIP,serverPort,clientPort
	' response : 0

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	Dim serverName,serverIP,clientIP,serverPort,clientPort
	If UBound(parameters)<4 Then Call Error("not enough parameter")
	'serverName	= REQUESTERID
	serverName	= Parameters(0)
	serverIP	= Parameters(1)
	clientIP	= Parameters(2)
	serverPort	= Parameters(3)
	clientPort	= Parameters(4)
	If Not (IsNumeric(serverPort) Or IsNumeric(clientPort)) Then ' should be numeric
		Call Error("invalid parameter")
	End If
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_StartAchievementServer"
		Call .InitCommand()
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
		Call .AppendParam("@serverIP",adVarChar,adParamInput,4000,serverIP)
		Call .AppendParam("@clientIP",adVarChar,adParamInput,4000,clientIP)
		Call .AppendParam("@serverPort",adInteger,adParamInput,,serverPort)
		Call .AppendParam("@clientPort",adInteger,adParamInput,,clientPort)
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

	Call Ok(0)
%>

