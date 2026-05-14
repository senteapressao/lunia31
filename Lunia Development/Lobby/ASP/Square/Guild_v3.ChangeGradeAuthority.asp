<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : guildId,grade,guildMemberId,authority
	' response : grade,authority
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim guildId,grade,guildMemberId,authority
	If UBound(parameters)<3 Then Call Error("not enough parameter")
	guildId = Parameters(0)
	grade = Parameters(1)
	guildMemberId = Parameters(2)
	authority = Parameters(3)
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_ChangeGradeAuthority"
		Call .InitCommand()
		Call .AppendParam("@guildId",adInteger,adParamInput,,guildId)
		Call .AppendParam("@guildMemberId",adInteger,adParamInput,,guildMemberId)		
		Call .AppendParam("@grade",adUnsignedTinyint,adParamInput,,grade)		
		Call .AppendParam("@authority",adUnsignedTinyint,adParamInput,,authority)		
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

	Call Ok(grade & SEPARATOR & authority)
%>