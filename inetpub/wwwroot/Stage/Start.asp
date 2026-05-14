<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : (POST)
	'            address,port,number of rooms,service stage group hash list(subseparated stage group hash,capacity)
	' response : nothing
	
	' ** remarks **
	' actually,stage group has list sould be separated by SEPARATOR than SUBSEPARATOR
	'	coz it's last parameters though the number of element is variable.
	'	for compatibility with last protocol version(there is only stage group hash without capacity)
	'	i made it to be separated by SUBSEPARATOR
	Dim rowData
	'rowData=getString(Request.BinaryRead(Request.TotalBytes))
	rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	
	' preparing request parameters
	parameters=Split(rowData,SEPARATOR)
	Dim serverName,address,port,rooms,isDevOnly,serverLocation,stageGroupHashes
	If UBound(Parameters)<4 Then Call Error("not enough parameter")
	serverName = REQUESTERID
	address = Parameters(0)
	port = Parameters(1)
	rooms = Parameters(2)
	isDevOnly = Parameters(3)
	serverLocation = Parameters(4)
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_StartStageServer"
		Call .InitCommand()
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
		Call .AppendParam("@address",adVarChar,adParamInput,4000,address)
		Call .AppendParam("@port",adInteger,adParamInput,,port)
		Call .AppendParam("@roomCount",adInteger,adParamInput,,rooms)
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
	set conn = Server.CreateObject("ADODB.Connection")

	Dim row,stageGroupHash,capacity
	For i=5 To Ubound(Parameters)
		row = Split(Parameters(i),SUBSEPARATOR)
		stageGroupHash = row(0)
		capacity = "default"
		If Ubound(row)>0 Then capacity = row(1)

		sql = sql &"insert into StageGroups(serverName,stageGroupHash,capacity)"&_
			" values ("& SqlQuot(serverName) &","& stageGroupHash &","& capacity &");"
	Next
	conn.Open(stageDBconnectionString)
	conn.Execute(sql)
	conn.Close()

	Call Ok("Ok")
%>