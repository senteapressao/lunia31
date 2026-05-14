<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : guildName,guildAlias,masterName,message,gradeName0,gradeName1,gradeName2,gradeName3,gradeName4
	' response : guildId
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim guildName,guildAlias,masterName,message,gradeName0,gradeName1,gradeName2,gradeName3,gradeName4
	If UBound(parameters)<8 Then Call Error("not enough parameter")
	guildName = Parameters(0)
	guildAlias = Parameters(1)
	masterName = Parameters(2)
	message = Parameters(3)
	gradeName0 = Parameters(4)
	gradeName1 = Parameters(5)
	gradeName2 = Parameters(6)
	gradeName3 = Parameters(7)
	gradeName4 = Parameters(8)
	
	Dim guildId
	
	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Create"
		Call .InitCommand()
		Call .AppendParam("@guildName",adVarWChar,adParamInput,14,guildName)
		Call .AppendParam("@guildAlias",adVarWChar,adParamInput,24,guildAlias)
		Call .AppendParam("@masterName",adVarWChar,adParamInput,50,masterName)
		Call .AppendParam("@message",adVarWChar,adParamInput,100,message)
		Call .AppendParam("@gradeName0",adVarWChar,adParamInput,20,gradeName0)
		Call .AppendParam("@gradeName1",adVarWChar,adParamInput,20,gradeName1)
		Call .AppendParam("@gradeName2",adVarWChar,adParamInput,20,gradeName2)
		Call .AppendParam("@gradeName3",adVarWChar,adParamInput,20,gradeName3)
		Call .AppendParam("@gradeName4",adVarWChar,adParamInput,20,gradeName4)
		Call .AppendParam("@guildId",adInteger,adParamInputOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		guildId = sph.GetParamValue("@guildId")
	End If
	Set sph = Nothing

	Call Ok(guildId & SEPARATOR & guildName)
%>