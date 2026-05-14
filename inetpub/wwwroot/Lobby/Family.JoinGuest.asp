<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : accountName,characterName
	' response : familyId,familyMemberId
	
	Dim i,j,k
	Dim sph,rs,rs2,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim accountName,characterName
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	accountName = Parameters(0)
	characterName = Parameters(1)
	
	Dim familyId,familyMemberId
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Family_JoinByGuest"
		Call .InitCommand()
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@familyId",adInteger,adParamOutput,,null)
		Call .AppendParam("@familyMemberId",adInteger,adParamOutput,,null)
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
		familyMemberId = sph.GetParamValue("@familyMemberId")
	End If
	Set sph = Nothing
	
	Call Ok(familyId & SEPARATOR & familyMemberId)
%>