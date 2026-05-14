<!--#include file="../common.asp"-->
<!--#include file="../publisher.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : accountName,password,failCount,lockExpired
	' response : 

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim accountName,password,failCount,lockExpired
	If UBound(parameters)<3 Then Call Error("not enough parameter")
	accountName	= Parameters(0)
	password	= Parameters(1)
	failCount	= Parameters(2)
	lockExpired	= Parameters(3)
	
	' call Stored procedure
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.Auth_2nd_Sign"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		Call .AppendParam("@password",adVarWChar,adParamInput,50,password)
		Call .AppendParam("@failCount",adSmallInt,adParamInput,,failCount)
		Call .AppendParam("@lockExpired",adVarWChar,adParamInput,50,lockExpired)
		blsResult = .ExecNoRecords()
	End with

	Dim ret
	ret = sphn.GetParamValue("RETURN_VALUE")
	if ret<>0 Then
		Call Error(ret)
	End If
	set sphn = Nothing

	Call Ok(0)
%>