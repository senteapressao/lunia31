<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : character name, stage group hash, access level of the stage group
	' response : stage server address, port, keycode

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim characterName,stageGroupHash,accessLevel,step
	Dim	address,port,reservationSerial
	If UBound(parameters)<3 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	stageGroupHash = Parameters(1)
	accessLevel = Parameters(2)
	step = Parameters(3)
	
	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_QuickJoinStage"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@stageGroupHash",adInteger,adParamInput,,stageGroupHash)
		Call .AppendParam("@accessLevel",adInteger,adParamInput,,accessLevel)
		Call .AppendParam("@step",adTinyInt,adParamInput,,step)
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