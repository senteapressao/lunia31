<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,familyMemberId
	' response : 0
	
	Dim i,j,k
	Dim sph,rs,rs2,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName,familyMemberId,characterName_target,familyMemberId_target
	If UBound(parameters)<3 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	familyMemberId = Parameters(1)
	characterName_target = Parameters(2)
	familyMemberId_target = Parameters(3)
	
	Dim familyId,isGuest
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Family_Kick"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@familyMemberId",adInteger,adParamInput,,familyMemberId)
		Call .AppendParam("@characterName_target",adVarWChar,adParamInput,50,characterName_target)
		Call .AppendParam("@familyMemberId_target",adInteger,adParamInput,,familyMemberId_target)
		Call .AppendParam("@familyId",adInteger,adParamOutput,,null)
		Call .AppendParam("@isGuest",adBoolean,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		familyId = sph.GetParamValue("@familyId")
		isGuest = sph.GetParamValue("@isGuest")
		isGuest = Abs(CInt(isGuest))
	End If
	Set sph = Nothing

	Call Ok(characterName_target & SEPARATOR & familyId & SEPARATOR & familyMemberId_target & SEPARATOR & isGuest)
%>