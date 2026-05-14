<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request	: characterSerial
	' response	: {achievementHash(complete,logDate,isShow,isSend)}
	'			,{achievementHash(woring)}
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim rs2
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterSerial
	If UBound(parameters)<0 Then Call Error("not enough parameter")	
	characterSerial	= Parameters(0)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.achievementGet"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterSerial",adBigInt,adParamInput,,characterSerial)
		blsResult = .ExecRecordset()
		'blsResult = .ExecRecordset2()
	End with

	If blsResult Then
		Set rs = sphn.rs
		'Set rs2 = sphn.rs2
	Else
		Call Error(1)
	End If
	set sphn = Nothing
	
	Do Until rs.Eof
		retString = retString & rs(0) & SUBSEPARATOR & FormatDt(rs(1),"SQL_TM") & SUBSEPARATOR & rs(2) & SUBSEPARATOR & rs(3)
		rs.MoveNext
		If Not rs.Eof Then retString = retString & SUBSEPARATOR
	Loop
	
	retString = retString & SEPARATOR

	'Working
	'Do Until rs2.Eof
	'	retString = retString & rs2(0)
	'	rs2.MoveNext
	'	If Not rs.Eof Then retString = retString & SUBSEPARATOR
	'Loop

	Call Ok(retString)
%>