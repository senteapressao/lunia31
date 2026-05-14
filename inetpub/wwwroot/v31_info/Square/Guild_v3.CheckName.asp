<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : guildName
	' response : 0
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim guildName
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	guildName = Parameters(0)
	
	Dim isExists
	
	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_CheckName"
		Call .InitCommand()
		Call .AppendParam("@guildName",adVarWChar,adParamInput,14,guildName)
		Call .AppendParam("@isExists",adBoolean,adParamOutput,,isExists)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		isExists = sph.GetParamValue("@isExists")
	End If
	Set sph = Nothing

	If isExists Then
		Call Ok(1)
	Else
		Call Ok(0)
	End If
%>
