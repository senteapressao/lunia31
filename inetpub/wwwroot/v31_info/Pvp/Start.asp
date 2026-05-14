<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  :  address,port,minLevel,maxLevel
	' response : nothing

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	
	' preparing request parameters	
	Dim channelName,address,port,minLevel,maxLevel,isBalanced	
	channelName = REQUESTERID		
	address = Parameters(0)
	port = Parameters(1)
	minLevel = Parameters(2)
	maxLevel = Parameters(3)
	isBalanced = Parameters(4)
	
	' check parameter validation
	If Not(IsNumeric(port) AND IsNumeric(minLevel) AND IsNumeric(maxLevel)) Then ' should be numeric
		Call Error("invalid parameter")
	End If
	
	' sql query to retrieve Pvp
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_StartPvpServer"
		Call .InitCommand()
		Call .AppendParam("@channelName",adVarWChar,adParamInput,50,channelName)
		Call .AppendParam("@address",adVarChar,adParamInput,50,address)
		Call .AppendParam("@port",adInteger,adParamInput,,port)
		Call .AppendParam("@minLevel",adInteger,adParamInput,,minLevel)
		Call .AppendParam("@maxLevel",adInteger,adParamInput,,maxLevel)
		Call .AppendParam("@isBalanced",adBoolean,adParamInput,,isBalanced)
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

