<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	' response : ladderPoint,dailyPlayCount,playCount,winCount

	Dim i,j,k
	Dim sphn,blsResult : blsResult = False
	Dim retString : retString = ""
	
	' preparing request parameters	
	Dim characterName
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	
	Set sphn = new SPHelper_NoTran
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.LadderGetPoint"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset()
	End with

	If blsResult Then
		Set rs = sphn.rs
	End If
	set sphn = Nothing
	
	If rs.Eof Then Call Error("no recordset")
	
	If Not(rs.Eof) Then
		'retString = retString & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & rs(3)
		retString = retString & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & rs(3)
	Else
		'retString = retString & 1500 & SEPARATOR & 0 & SEPARATOR & 0 & SEPARATOR & 0
		retString = retString & 1500 & SEPARATOR & 0 & SEPARATOR & 0 & SEPARATOR & 0
	End If

	Call Ok("Ok")
%>