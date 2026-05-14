<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,comment,timeout
	' response : 0

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim characterName,comment,timeout

	If UBound(Parameters)<1 Then Call Error("not enough parameter")	
	characterName = Parameters(0)
	comment = Parameters(1)
	If UBound(Parameters)>1 Then timeout=CLng(Parameters(2)) Else timeout=60*60*24 '1 day

	comment = "(" & remoteIp & ")banned:" & comment

	' call Stored procedure
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.Block"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@timeoutSeconds",adInteger,adParamInput,,timeout)
		Call .AppendParam("@comment",adVarChar,adParamInput,8000,comment)
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@remoteIp",adVarWChar,adParamInput,50,null)
		blsResult = .ExecNoRecords()
	End with

	Dim ret
	ret = sphn.GetParamValue("RETURN_VALUE")
	if ret<>0 Then
		Call Error(ret)
	End If
	set sphn = Nothing

	Call Ok(ret)
%>