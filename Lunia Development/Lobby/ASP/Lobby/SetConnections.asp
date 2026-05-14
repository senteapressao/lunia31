<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : connectionCount
	' response : connectionCount

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim connectionCount
	
	If UBound(Parameters)<0 Then Call Error("not enough parameter")
	connectionCount = Parameters(0)

	' call Stored procedure
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = gateDBconnectionString
		.SPName = "dbo.SetConnections"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@serverCode",adSmallInt,adParamInput,,ServerGroupCode)
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,REQUESTERID)
		Call .AppendParam("@connectionCount",adInteger,adParamInput,,connectionCount)
		blsResult = .ExecNoRecords()
	End with

	Dim ret
	ret = sphn.GetParamValue("RETURN_VALUE")
	if ret<>0 Then
		Call Error(ret)
	End If
	set sphn = Nothing

	Call Ok(connectionCount)
%>
