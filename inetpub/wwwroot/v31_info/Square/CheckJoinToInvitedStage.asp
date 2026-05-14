<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : InviterName, stageGroupHash, accessLevel
	' response : only result (exist or non-exist)
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim InviterName, stageGroupHash, accessLevel
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	InviterName		= Parameters(0)
	stageGroupHash	= Parameters(1)
	accessLevel		= Parameters(2)
	
	' TODO: parameter validation

	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_CheckJoinToInvitedStage"
		Call .InitCommand()
		Call .AppendParam("@InviterName",adVarWChar,adParamInput,50, inviterName)
		Call .AppendParam("@stageGroupHash",adInteger,adParamInput,,stageGroupHash)
		Call .AppendParam("@accessLevel",adInteger,adParamInput,,accessLevel)
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

	Call Ok("0")
%>
