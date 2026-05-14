<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : address,port,capacity
	' response : 
		
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False

	Dim serverName,address,port,capacity
	serverName = REQUESTERID
	address = Parameters(0)
	port = Parameters(1)
	capacity = Parameters(2)
	
	If IsEmpty(address) OR address="" Then address="(any)"
	
	' check parameter validation
	If Not (IsNumeric(port) OR IsNumeric(capacity)) Then ' should be numeric
		Call Error("invalid parameter")
	End If
	
	' call Stored procedure	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.StartLobbyServer"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
		Call .AppendParam("@address",adVarChar,adParamInput,50,address)
		Call .AppendParam("@port",adInteger,adParamInput,50,port)
		Call .AppendParam("@capacity",adInteger,adParamInput,50,capacity)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
	End If
	set sphn = Nothing

	Call Ok("Ok")
%>