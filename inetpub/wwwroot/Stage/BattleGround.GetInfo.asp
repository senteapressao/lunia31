<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : temporary key, password, charactername, teamNumber
	' response : temporary key, stage server address, port, key code

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	Dim serverName,roomNumber
	Dim	battleGroundKillCount,battleGroundLimitTime
	serverName = REQUESTERID
	roomNumber = parameters(0)

	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_GetBattleGroundInfo"
		Call .InitCommand()
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
		Call .AppendParam("@roomNumber",adInteger,adParamInput,,roomNumber)
		Call .AppendParam("@battleGroundKillCount",adInteger,adParamOutput,,null)
		Call .AppendParam("@battleGroundLimitTime",adSingle,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		battleGroundKillCount = sph.GetParamValue("@battleGroundKillCount")
		battleGroundLimitTime = sph.GetParamValue("@battleGroundLimitTime")
	End If
	set sph = Nothing

	retString = battleGroundKillCount & SEPARATOR & battleGroundLimitTime
	
	Call Ok(retString)
%>