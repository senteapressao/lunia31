<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	'Parameters = Split("Teles47449731#20570475%3cCS%3e20210620173235", SEPARATOR)
	' request  : character name, stage group hash, access level of the stage group, stage name(key)
	' response : stage server address, port, stage serial

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim characterName,stageGroupHash,accessLevel,step,stageName
	Dim	address,port,reservationSerial
	
	'req << sender->GetCharacterName()
	'<< packet.Stage.StageGroupHash
	'<< packet.Stage.Level
	'<< packet.Stage.Difficulty
	'<< packet.Password;
	
	
	If UBound(parameters)<3 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	stageGroupHash = Parameters(1)
	accessLevel = Parameters(2)
	step = Parameters(3)
	If UBound(Parameters)>3 Then stageName = Parameters(4) Else stageName = ""

	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = false
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_NewStage"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@stageGroupHash",adInteger,adParamInput,,stageGroupHash)
		Call .AppendParam("@accessLevel",adInteger,adParamInput,,accessLevel)
		Call .AppendParam("@step",adTinyInt,adParamInput,,step)
		Call .AppendParam("@password",adVarWChar,adParamInput,80,stageName)
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