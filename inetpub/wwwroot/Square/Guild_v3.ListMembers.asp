<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : guildId
	' response : characterName,grade,datediff(n,a.latestLogout,getdate()),classNumber,stageLevel,pvpLevel,contributed
	
	Dim i,j,k
	Dim sph,rs,rs2,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim guildId
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	guildId = Parameters(0)
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_ListMembers"
		Call .InitCommand()
		Call .AppendParam("@guildId",adInteger,adParamInput,,guildId)
		blsResult = .ExecRecordset()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		Set rs = sph.rs
	End If
	Set sph = Nothing
	
	i = 0
	Do Until rs.Eof
		If i>0 Then retString = retString & SEPARATOR
		retString = retString & rs(0) & SUBSEPARATOR & rs(1) &_
				SUBSEPARATOR & rs(2) & SUBSEPARATOR & rs(3) &_
				SUBSEPARATOR & rs(4) & SUBSEPARATOR & rs(5) &_
				SUBSEPARATOR & rs(6) & SUBSEPARATOR & rs(7)
				
		rs.MoveNext : i=i+1
	Loop
	rs.close
	
	Call Ok(retString)
%>