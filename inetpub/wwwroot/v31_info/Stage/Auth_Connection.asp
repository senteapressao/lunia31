<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()	
	' request  : reservationSerial,tempId
	' response : tempId
	'			,characterName	
	'			,roomNumber,stageGroupHash,accessLevel,step,password,capacity,teamNumber
	'			,stageStates,lastPlayedStageGroupHash,lastPlayedStageLevel,uniqueKey
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim serverName,reservationSerial,tempId
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	serverName = REQUESTERID
	reservationSerial = Parameters(0)
	tempId = Parameters(1)
	
	Dim characterName
	Dim roomNumber,stageGroupHash,accessLevel,step,password,capacity,teamNumber,stageStates,lastPlayedStageGroupHash,lastPlayedStageLevel,uniqueKey
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_AuthConnection"
		Call .InitCommand()
		Call .AppendParam("@reservationSerial",adBigInt,adParamInput,,reservationSerial)
		Call .AppendParam("@severName",adVarWChar,adParamInput,50,serverName)
		Call .AppendParam("@characterName",adVarWChar,adParamOutput,50,null)
		Call .AppendParam("@roomNumber",adInteger,adParamOutput,,null)
		Call .AppendParam("@stageGroupHash",adInteger,adParamOutput,,null)
		Call .AppendParam("@accessLevel",adInteger,adParamOutput,,null)
		Call .AppendParam("@step",adTinyInt,adParamOutput,,null)
		Call .AppendParam("@password",adVarWChar,adParamOutput,80,null)
		Call .AppendParam("@capacity",adInteger,adParamOutput,,null)
		Call .AppendParam("@teamNumber",adInteger,adParamOutput,,null)
		Call .AppendParam("@stageStates",adInteger,adParamOutput,,0)
		Call .AppendParam("@characterStates",adInteger,adParamOutput,,0)
		Call .AppendParam("@lastPlayedStageGroupHash",adInteger,adParamOutput,,null)
		Call .AppendParam("@lastPlayedStageLevel",adSmallInt,adParamOutput,,null)
		Call .AppendParam("@uniqueKey",adBigInt,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		characterName = sph.GetParamValue("@characterName")
		roomNumber = sph.GetParamValue("@roomNumber")
		stageGroupHash = sph.GetParamValue("@stageGroupHash")
		accessLevel = sph.GetParamValue("@accessLevel")
		step = sph.GetParamValue("@step")
		password = sph.GetParamValue("@password")
		teamNumber = sph.GetParamValue("@teamNumber")
		capacity = sph.GetParamValue("@capacity")
		stageStates = sph.GetParamValue("@stageStates")
		'characterStates = sph.GetParamValue("@characterStates")
		lastPlayedStageGroupHash = sph.GetParamValue("@lastPlayedStageGroupHash")
		lastPlayedStageLevel = sph.GetParamValue("@lastPlayedStageLevel")
		uniqueKey = sph.GetParamValue("@uniqueKey")
	End If
	set sph = Nothing

	Response.Write tempId &_
			SEPARATOR & characterName &_
			SEPARATOR & roomNumber & SEPARATOR & stageGroupHash & SEPARATOR & accessLevel & SEPARATOR & step &_
			SEPARATOR & password & SEPARATOR & capacity & SEPARATOR & teamNumber &_
			SEPARATOR & stageStates & SEPARATOR & lastPlayedStageGroupHash & SEPARATOR & lastPlayedStageLevel &_
			SEPARATOR & uniqueKey
%>