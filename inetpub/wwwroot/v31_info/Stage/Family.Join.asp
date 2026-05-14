<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,familyId,memorialDay
	' response : familyId,familyMemberId
	
	Dim i,j,k
	Dim sph,rs,rs2,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName,familyId,memorialDay
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	familyId = Parameters(1)
	memorialDay = Parameters(2)
	
	Dim familyMemberId
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Family_Join"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@familyId",adInteger,adParamInputOutput,,familyId)
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
	
	' update memorialDay default
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_FamilyMember_Update"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@familyMemberId",adInteger,adParamInput,,familyMemberId)
		Call .AppendParam("@expire_gift1",adDBTimeStamp,adParamInput,,null)
		Call .AppendParam("@memorialDay",adDBTimeStamp,adParamInput,,FormatDt(memorialDay,"SQL_TM"))
		blsResult = .ExecNoRecords()
	End with
	
	Call Ok(familyId & SEPARATOR & familyMemberId)
%>