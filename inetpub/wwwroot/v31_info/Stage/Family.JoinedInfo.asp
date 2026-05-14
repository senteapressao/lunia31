<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	' response : familyId,createDate,familyMemberId,isGuest,isOnline,playTime,expire_gift1,memorialDay,joinedDate
	'			,{familyMemberId,characterName,classsNumber,stageLevel,pvpLevel,isGuest,isOnline,playTime,joinedDate,lastLoggedDate}

	Dim i,j,k
	Dim sph,rs,rs2,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	characterName = Parameters(0)

	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_Family_JoinedInfo"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset2()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		Set rs = sph.rs
		Set rs2 = sph.rs2
	End If
	Set sph = Nothing

	If Not(rs.Eof) Then
		retString = retString & rs(0) &_
				SEPARATOR & FormatDt(rs(1),"SQL_TM") & SEPARATOR & rs(2) &_
				SEPARATOR & rs(7) & SEPARATOR & rs(8) &_
				SEPARATOR & rs(9) & SEPARATOR & FormatDt(rs(10),"SQL_TM") &_
				SEPARATOR & FormatDt(rs(11),"SQL") & SEPARATOR & FormatDt(rs(12),"SQL_TM")
	End If
	rs.close

	retString = retString & SEPARATOR

	i = 0
	Do Until rs2.Eof
		If i>0 Then retString = retString & SUBSEPARATOR
		retString = retString & rs2(0) &_
			SUBSEPARATOR & rs2(1) & SUBSEPARATOR & rs2(2) &_
			SUBSEPARATOR & rs2(3) & SUBSEPARATOR & rs2(4) &_
			SUBSEPARATOR & rs2(5) & SUBSEPARATOR & rs2(6) &_
			SUBSEPARATOR & rs2(7) & SUBSEPARATOR & FormatDt(rs2(8),"SQL_TM") &_
			SUBSEPARATOR & FormatDt(rs2(9),"SQL_TM")

		rs2.MoveNext : i=i+1
	Loop
	rs2.Close

	Call Ok(retString)
%>