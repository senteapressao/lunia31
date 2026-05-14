<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName1,characterName2,memorialDay
	' response : familyId,familyMemberId1,familyMemberId1
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim characterName1,characterName2,memorialDay
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	characterName1 = Parameters(0)
	characterName2 = Parameters(1)
	memorialDay = Parameters(2)
	
	Dim familyId,familyMemberId1,familyMemberId2
	
	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Family_Create"
		Call .InitCommand()
		Call .AppendParam("@characterName1",adVarWChar,adParamInput,50,characterName1)
		Call .AppendParam("@characterName2",adVarWChar,adParamInput,50,characterName2)
		Call .AppendParam("@familyId",adInteger,adParamInputOutput,,null)
		Call .AppendParam("@familyMemberId1",adInteger,adParamInputOutput,,null)
		Call .AppendParam("@familyMemberId2",adInteger,adParamInputOutput,,null)
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
		familyMemberId1 = sph.GetParamValue("@familyMemberId1")
		familyMemberId2 = sph.GetParamValue("@familyMemberId2")
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
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName1)
		Call .AppendParam("@familyMemberId",adInteger,adParamInput,,familyMemberId1)
		Call .AppendParam("@expire_gift1",adDBTimeStamp,adParamInput,,null)
		Call .AppendParam("@memorialDay",adDBTimeStamp,adParamInput,,FormatDt(memorialDay,"SQL"))
		blsResult = .ExecNoRecords()

		Call .TerminateCommand()
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName2)
		Call .AppendParam("@familyMemberId",adInteger,adParamInput,,familyMemberId2)
		Call .AppendParam("@expire_gift1",adDBTimeStamp,adParamInput,,null)
		Call .AppendParam("@memorialDay",adDBTimeStamp,adParamInput,,FormatDt(memorialDay,"SQL"))
		blsResult = .ExecNoRecords()
	End with
	
	Call Ok(familyId & SEPARATOR & familyMemberId1 & SEPARATOR & familyMemberId2)
%>