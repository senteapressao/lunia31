<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : guildAlias
	' response : 0
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim guildAlias
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	guildAlias = Parameters(0)
	
	Dim isExists
	
	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_CheckAlias"
		Call .InitCommand()
		Call .AppendParam("@guildAlias",adVarWChar,adParamInput,24,guildAlias)
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
