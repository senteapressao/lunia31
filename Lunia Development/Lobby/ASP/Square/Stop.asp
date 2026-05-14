<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : nothing
	' response : nothing

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	
	' preparing request parameters
	Dim serverName
	If UBound(Parameters)=0 Then serverName = Parameters(0)
	If serverName="" Then serverName = REQUESTERID

	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_StopSquareServer"
		Call .InitCommand()
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
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