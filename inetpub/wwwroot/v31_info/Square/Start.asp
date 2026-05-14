<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : address,port,roomCount
	'			,squareinfo(subseparated: squareName,stageGroupHash,accessLevel,capacity)...
	' response : 0

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	Dim serverName,address,port,roomCount,squareInfo
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	serverName = REQUESTERID
	address = Parameters(0)
	port = Parameters(1)
	roomCount = Parameters(2)
	
	If Not (IsNumeric(port) OR IsNumeric(roomCount)) Then ' should be numeric
		Call Error("invalid parameter")
	End If
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_StartSquareServer"
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
	Dim squareName,stageGroupHash,accessLevel,capacity,orderNumber
	Dim roomNumber : roomNumber = 0
	For i=3 To Ubound(Parameters)
		squareInfo = Parameters(i)
		Dim arr : arr = Split(squareInfo,SUBSEPARATOR)

		squareName = arr(0)
		stageGroupHash = arr(1)
		accessLevel = arr(2)
		capacity = arr(3)
		orderNumber = arr(4)
		
		sql = "insert dbo.SquareStages(serverName,roomNumber,capacity,stageGroupHash,accessLevel,squareName,orderNumber)"&_
			" select ?,?,?,?,?,?,?"
		with Command
			.ActiveConnection = stageDBconnectionString
			.CommandText = sql
			.CommandType = adCmdText
			.Parameters.Append .CreateParameter("@serverName",adVarWChar,adParamInput,50,serverName)
			.Parameters.Append .CreateParameter("@roomNumber",adInteger,adParamInput,,roomNumber)
			.Parameters.Append .CreateParameter("@capacity",adSmallInt,adParamInput,,capacity)
			.Parameters.Append .CreateParameter("@stageGroupHash",adInteger,adParamInput,,stageGroupHash)
			.Parameters.Append .CreateParameter("@accessLevel",adInteger,adParamInput,,accessLevel)
			.Parameters.Append .CreateParameter("@squareName",adVarWChar,adParamInput,50,squareName)
			.Parameters.Append .CreateParameter("@orderNumber",adSmallInt,adParamInput,,orderNumber)
			.Execute ,,adExecuteNoRecords
			Do Until Command.Parameters.Count = 0
				Command.Parameters.Delete (0)
			Loop
		End with

		roomNumber = roomNumber + 1
	Next

	Call Ok(0)
%>

