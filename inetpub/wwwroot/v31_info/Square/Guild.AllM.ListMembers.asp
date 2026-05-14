<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	' response : Friend list (separated : characterName, class, level, isFriend )
	
	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim rs2,rs3
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim guildId
	guildId=Parameters(0)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.GuildListMembers"
		Call .InitCommand()
		Call .AppendParam("@guildId",adInteger,adParamInput,,guildId)
		blsResult = .ExecRecordset()
	End with
	
	If blsResult Then
		Set rs = sphn.rs
	End If
	set sphn = Nothing
	
	i=0
	Do Until rs.Eof
		If i>0 Then
				retString = retString & SEPARATOR
		End If
			
		retString = retString & rs(0) & SUBSEPARATOR & rs(1) &_
				SUBSEPARATOR & rs(2) & SUBSEPARATOR & rs(3) &_
				SUBSEPARATOR & rs(4) & SUBSEPARATOR & rs(5)

		rs.MoveNext:i=i+1
	Loop
	rs.close
	
	Call OK(retString)
%>
