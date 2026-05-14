<!--#include file="../common.asp"-->
<!--#include file="../publisher.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	' request  : account name, password, ip, hwid
	' response : account name
	Call Init()
	If UBound(parameters)<2 Then Call Error("not enough parameter") 

	'Dim rowData
	'rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	
	' preparing request parameters
	Dim accountName, password, hwid
	
	accountName=Parameters(0)
	password=Parameters(1)
	remoteIp=Parameters(2)
	hwid=Parameters(3)

	' check parameter validation
	' call Stored procedure
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.AuthAccount"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,REQUESTERID)
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		Call .AppendParam("@password",adVarWChar,adParamInput,50,password)
		Call .AppendParam("@remoteIp",adVarWChar,adParamInput,50,remoteIp)
		Call .AppendParam("@hwid",adVarWChar,adParamInput,50,hwid)
		blsResult = .ExecNoRecords()
	End with

	Dim ret
	ret = sphn.GetParamValue("RETURN_VALUE")
	if ret=0 Then
		Call Ok(accountName & SEPARATOR & 0)
	Else
		Call Error(ret)
	End If
	set sphn = Nothing
%>