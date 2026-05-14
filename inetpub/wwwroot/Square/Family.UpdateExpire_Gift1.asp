<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,familyMemberId,expire_gift1
	' response : 
	
	Dim i,j,k
	Dim sph,rs,rs2,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName,familyMemberId,expire_gift1
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	familyMemberId = Parameters(1)
	expire_gift1 = Parameters(2)
	
	' update expire_gift1 default
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_FamilyMember_Update"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@familyMemberId",adInteger,adParamInput,,familyMemberId)
		Call .AppendParam("@expire_gift1",adDBTimeStamp,adParamInput,,FormatDt(expire_gift1,"SQL"))
		Call .AppendParam("@memorialDay",adDBTimeStamp,adParamInput,,null)
		blsResult = .ExecNoRecords()
	End with
	
	Call Ok(0)
%>