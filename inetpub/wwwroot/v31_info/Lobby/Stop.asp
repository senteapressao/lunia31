<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : address,port,capacity
	' response : 
		
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False

	Dim serverName
	If UBound(Parameters)=0 Then serverName = Parameters(0)
	If serverName="" Then serverName = REQUESTERID
	
	' call Stored procedure	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.StopLobbyServer"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
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
	
	Dim conn,sql
	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open(characterDBconnectionString)
	sql = "select count('x') from dbo.LobbyConnections (nolock)"
	If conn.Execute(sql)(0)=0 Then
		conn.Close()
		conn.Open(guildDBconnectionString)
		sql = "update dbo.familyMember set isOnline=0"
		conn.Execute(sql)				
	End If
	conn.Close()
	
	Call Ok("Ok")
%>
