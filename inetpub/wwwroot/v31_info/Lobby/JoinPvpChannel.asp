<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : character name, channelName
	' response : pvp channel address, port, key
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	
	' preparing request parameters
	Dim characterName,channelName
	Dim address,port,reservationSerial
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	channelName = Parameters(1)
	
	' TODO: parameter validation

	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_JoinPvpChannel"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@channelName",adVarWChar,adParamInput,50,channelName)
		Call .AppendParam("@address",adVarChar,adParamOutput,50,null)
		Call .AppendParam("@port",adInteger,adParamOutput,,null)
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
		address = sph.GetParamValue("@address")
		port = sph.GetParamValue("@port")
		reservationSerial = sph.GetParamValue("@reservationSerial")
	End If
	set sph = Nothing

	Call Ok(address & SEPARATOR & port & SEPARATOR & reservationSerial)
%>