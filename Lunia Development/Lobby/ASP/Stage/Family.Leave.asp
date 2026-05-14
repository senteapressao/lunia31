<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,familyMemberId,isPenalty
	' response : characterName,familyId,familyMemberId,isGuest
	
	Dim i,j,k
	Dim sph,rs,rs2,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName,familyMemberId,isPenalty
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	familyMemberId = Parameters(1)
	isPenalty = Parameters(2)
	
	Dim familyId,isGuest
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Family_Leave"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@familyMemberId",adInteger,adParamInput,,familyMemberId)
		Call .AppendParam("@isPenalty",adUnsignedTinyInt,adParamInput,,isPenalty)
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

	Call Ok(characterName & SEPARATOR & familyId & SEPARATOR & familyMemberId & SEPARATOR & isGuest)
%>