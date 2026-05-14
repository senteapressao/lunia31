<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request : characterName,playTime
	' response : 0
		
	Dim sph,blsResult : blsResult = False

	' preparing request parameters
	Dim characterName,playTime
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	playTime = Parameters(1)
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Family_Terminate"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@playTime",adInteger,adParamInput,,playTime)
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

	Call Ok(0)
%>