<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : (POST)address,port,roomCount
	'			,stageGroups(subseparated : stageGroupHash,capacity)
	' response : 0
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	Dim rowData
	rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)
	
	' preparing request parameters
	Dim serverName,address,port,roomCount,stageGroups
	If UBound(Parameters)<2 Then Call Error("not enough parameter")
	serverName = REQUESTERID
	address = Parameters(0)
	port = Parameters(1)
	roomCount = Parameters(2)

	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_StartStageServer"
		Call .InitCommand()
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
		Call .AppendParam("@address",adVarChar,adParamInput,50,address)
		Call .AppendParam("@port",adInteger,adParamInput,,port)
		Call .AppendParam("@roomCount",adInteger,adParamInput,,roomCount)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	End If
	set sph = Nothing

	Dim sql
	Dim stageGroupHash,capacity
	For i=3 To Ubound(Parameters)
		stageGroups = Parameters(i)
		Dim arr : arr = Split(stageGroups,SUBSEPARATOR)

		stageGroupHash = arr(0)
		If Ubound(arr)>0 Then capacity = arr(1) Else capacity = 4

		sql = "insert into StageGroups(serverName,stageGroupHash,capacity)"&_
			" select ?,?,?"
		with Command
			.ActiveConnection = stageDBconnectionString
			.CommandText = sql
			.CommandType = adCmdText
			.Parameters.Append .CreateParameter("@serverName",adVarWChar,adParamInput,50,serverName)
			.Parameters.Append .CreateParameter("@stageGroupHash",adInteger,adParamInput,,stageGroupHash)
			.Parameters.Append .CreateParameter("@capacity",adSmallInt,adParamInput,,capacity)
			.Execute ,,adExecuteNoRecords
			Do Until Command.Parameters.Count = 0
				Command.Parameters.Delete (0)
			Loop
		End with
	Next

	Call Ok(0)
%>