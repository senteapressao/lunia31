<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : mailId,characterName
	' response : characterName,sender,message,gameMoney,cashStampHash,isRead,status
	'			 itemHash,stackedCount,itemSerial,instance,itemExpire
	
	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim rs2
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim mailId,characterName
	If UBound(Parameters)<1 Then Call Error("not enough parameter")
	mailId = Parameters(0)
	characterName = Parameters(1)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.ReadMail"
		Call .InitCommand()
		Call .AppendParam("@mailId",adInteger,adParamInput,,mailId)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset2()
	End with
	
	If blsResult Then
		Set rs = sphn.rs
		Set rs2 = sphn.rs2
	Else
		Call Error(1)
	End If
	set sphn = Nothing

	If rs.Eof Then Call Error(1)
	retString = retString & rs(0) & SEPARATOR & rs(1) &_
				SEPARATOR & rs(2) & SEPARATOR & rs(3) &_
				SEPARATOR & rs(4) & SEPARATOR & rs(5) &_
				SEPARATOR & Abs(CInt(rs(6))) & SEPARATOR & rs(7)

	i=0
	Do Until rs2.Eof
		retString = retString & SEPARATOR & rs2(0) & SUBSEPARATOR & rs2(1) &_
				SUBSEPARATOR & rs2(2) & SUBSEPARATOR & rs2(3) & SUBSEPARATOR & FormatDt(rs2(4),"SQL_TM")

		rs2.MoveNext:i=i+1
	Loop
	rs.close
	
	Call OK(retString)
%>