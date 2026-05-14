<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	If UBound(parameters)<4 Then Call Error("not enough parameter")

	' request  : address,port,roomCount,isDevOnly
	'			, squareinfo(subseparated: squareName,stageGroupHash,accessLevel,capacity)...
	' response : 0

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	Dim serverName,address,port,roomCount,isDevOnly,serverLocation,squareInfo
	serverName = REQUESTERID
	address = Parameters(0)
	port = Parameters(1)
	roomCount = Parameters(2)
	isDevOnly = Parameters(3)
	serverLocation = Parameters(4)
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

	Dim conn,sql

	Dim squareName,stageGroupHash,accessLevel,capacity,orderNumber
	Dim roomNumber : roomNumber = 0
	For i=5 To Ubound(Parameters)
		Dim arr
		squareInfo = Parameters(i)
		arr = Split(squareInfo,SUBSEPARATOR)

		squareName = arr(0)
		stageGroupHash = arr(1)
		accessLevel = arr(2)
		capacity = arr(3)
		orderNumber = arr(4)

		sql = sql &"insert into SquareStages(serverName,roomNumber,capacity,stageGroupHash,accessLevel,squareName,orderNumber)"&_
			" values ("& SqlQuot(serverName) &","& roomNumber &","& capacity &","& stageGroupHash &","& accessLevel &","& SqlQuot(squareName) &","& orderNumber &");"

		roomNumber = roomNumber + 1
	Next

	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open(stageDBconnectionString)
	conn.Execute(sql)
	conn.Close()

	Call Ok(0)
%>

