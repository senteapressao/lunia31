<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	' response : guildId,guildName,guildAlias,message,masterName,memberCount,grade,guildMemberId,guildLevel,guildExp,contributed,playTime
	'			,guildPoint,shopStart,shopEnd,rank,expireDate,tax,taxPayDate
	'			gradeName,authority
	
	Dim i,j,k
	Dim sph,rs,rs2,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	
	Dim isBanned
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_JoinedInfo"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@isBanned",adBoolean,adParamOutput,,null)
		blsResult = .ExecRecordset2()
	End with

	If blsResult=False Then		
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		isBanned = sph.GetParamValue("@isBanned")
		Set rs = sph.rs
		Set rs2 = sph.rs2
	End If
	Set sph = Nothing
	
	If isBanned Then Call Ok("banned")
	
	i = 0
	Do Until rs.Eof
		If i>0 Then retString = retString & SEPARATOR
		retString = retString & rs(0) &_
				SEPARATOR & rs(1) & SEPARATOR & rs(2) &_
				SEPARATOR & rs(3) & SEPARATOR & rs(4) &_
				SEPARATOR & rs(5) & SEPARATOR & rs(6) &_
				SEPARATOR & rs(7) & SEPARATOR & rs(8) &_
				SEPARATOR & rs(9) & SEPARATOR & FormatDt(rs(10),"SQL_TM") &_
				SEPARATOR & rs(11) & SEPARATOR & rs(12) &_
				SEPARATOR & rs(13) & SEPARATOR & rs(14) &_
				SEPARATOR & rs(15) & SEPARATOR & FormatDt(rs(16),"SQL_TM") & SEPARATOR & FormatDt(rs(17),"SQL_TM") &_
				SEPARATOR & rs(18) & SEPARATOR & FormatDt(rs(19),"SQL_TM")
				
		rs.MoveNext : i=i+1
	Loop
	rs.close
	
	retString = retString & SEPARATOR
	
	i = 0
	Do Until rs2.Eof
		If i>0 Then retString = retString & SUBSEPARATOR
		retString = retString & rs2(0) & SUBSEPARATOR & rs2(1)
		
		rs2.MoveNext : i=i+1
	Loop
	rs2.close

	Call Ok(retString)
%>