<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : character name, stage group hash, access level of the stage group, stage name(key)
	' response : stage server address, port, stage serial

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim stageGroupHash,accessLevel,step,stageName,tempKey
	ReDim characters(Ubound(parameters)-5)
	Dim	serverName,roomNumber,address,port,reservationSerial
	
	If UBound(parameters)<4 Then Call Error("not enough parameter")
	stageGroupHash = Parameters(0)
	accessLevel = Parameters(1)
	step = Parameters(2)
	stageName = Parameters(3)
	tempKey = Parameters(4)
	j=0
	For i=5 To Ubound(parameters)
		If parameters(i)<>"" Then
			characters(j) = parameters(i)
			j=j+1
		End If
	Next

	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_ReserveStage"
		Call .InitCommand()
		Call .AppendParam("@stageGroupHash",adInteger,adParamInput,,stageGroupHash)
		Call .AppendParam("@accessLevel",adInteger,adParamInput,,accessLevel)
		Call .AppendParam("@step",adTinyInt,adParamInput,,step)
		Call .AppendParam("@password",adVarWChar,adParamInput,80,stageName)
		Call .AppendParam("@createSerial",adBigInt,adParamInput,,null)
		Call .AppendParam("@serverName",adVarWChar,adParamOutput,50,null)
		Call .AppendParam("@roomNumber",adInteger,adParamOutput,,null)
		Call .AppendParam("@address",adVarChar,adParamOutput,50,null)
		Call .AppendParam("@port",adInteger,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		serverName = sph.GetParamValue("@serverName")
		roomNumber = sph.GetParamValue("@roomNumber")
		address = sph.GetParamValue("@address")
		port = sph.GetParamValue("@port")
	End If
	set sph = Nothing

	retString = address & SEPARATOR & port

	For i=0 To Ubound(characters)
		Set sph = new SPHelper
		with sph
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = stageDBconnectionString
			.SPName = "dbo.public_JoinStageForReserve"
			Call .InitCommand()
			Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characters(i))
			Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
			Call .AppendParam("@roomNumber",adInteger,adParamInput,,roomNumber)
			Call .AppendParam("@reservationSerial",adBigInt,adParamOutput,,null)
			blsResult = .ExecNoRecords()
		End with

		If blsResult=False Then
			if sph.frk_n4ErrorCode>0 Then
				Call Error(sph.frk_n4ErrorCode)
			Else
				Call Error(sph.frk_strErrorText)
			End If
		Else
			reservationSerial = sph.GetParamValue("@reservationSerial")
		End If
		set sph = Nothing

		retString = retString & SEPARATOR & characters(i) & SEPARATOR & reservationSerial
	Next

	Call Ok(retString)
%>