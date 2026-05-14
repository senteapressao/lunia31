<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,familyMemberId,memorialDay
	' response : 
	
	Dim i,j,k
	Dim sph,rs,rs2,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName,familyMemberId,memorialDay
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	familyMemberId = Parameters(1)
	memorialDay = Parameters(2)
	
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
		Call .AppendParam("@memorialDay",adDBTimeStamp,adParamInput,,FormatDt(memorialDay,"SQL"))
		blsResult = .ExecNoRecords()
	End with
	
	Call Ok(0)
%>