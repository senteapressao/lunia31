<!--#include file="../common.asp"-->
<!--#include file="../publisher.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : accountName
	' response : isUse(0:not use/1:use),failCount,lockExpired

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim accountName
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	accountName	= Parameters(0)
	
	' call Stored procedure
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.Auth_2nd_Check"
		Call .InitCommand()
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		blsResult = .ExecRecordset()
	End with

	If blsResult Then
		Set rs = sphn.rs
	End If
	set sphn = Nothing
	
	If Not(rs.Eof) Then
		retString = "1" & SUBSEPARATOR & rs(0) & SUBSEPARATOR & FormatDt(rs(1),"SQL_TM")
	Else
		retString = "0" & SUBSEPARATOR &"0"& SUBSEPARATOR & "1900-01-01 00:00:00"
	End If

	Call Ok(retString)
%>