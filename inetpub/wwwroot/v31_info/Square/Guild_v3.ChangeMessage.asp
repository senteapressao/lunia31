<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : guildId,message,guildMemberId
	' response : message
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim guildId,message,guildMemberId
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	guildId = Parameters(0)
	message = Parameters(1)
	guildMemberId = Parameters(2)
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_ChangeMessage"
		Call .InitCommand()
		Call .AppendParam("@guildId",adInteger,adParamInput,,guildId)
		Call .AppendParam("@guildMemberId",adInteger,adParamInput,,guildMemberId)		
		Call .AppendParam("@message",adVarWChar,adParamInput,100,message)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	End If
	Set sph = Nothing

	Call Ok(message)
%>