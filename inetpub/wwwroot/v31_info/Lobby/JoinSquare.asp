<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : character name, square name, account name, oid
	' response : stage server address, port, stage group hash, access level of the stage group

	'Sub Pause(intSeconds)
	'	startTime = Time()
	'	Do Until DateDiff(“s”, startTime, Time(), 0, 0) > intSeconds
	'	Loop
	'End Sub

	Dim i,j,k
	Dim sph, sphn, rs,blsResult : blsResult = False

	' preparing request parameters
	Dim characterName,squareName
	Dim address,port,reservationSerial
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	squareName = Parameters(1)
	
	
	'BEGIN SYNTHETIC DELAY
	Function GetTextFromUrl(url)
	  Dim oXMLHTTP
	  Dim strStatusTest
	  Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	  oXMLHTTP.Open "GET", url, False
	  oXMLHTTP.Send
	  If oXMLHTTP.Status = 200 Then
		GetTextFromUrl = oXMLHTTP.responseText
		End If
	End Function
	'Dim sResult : sResult = GetTextFromUrl("http://localhost:81/sleep.php")
	
	'END SYNTHETIC DELAY
	
	' TODO: parameter validation
	
	' instantStateFlags
	' IsBalanced,IsSpectator off
	' IsPetNotUsable,IsEquipSetRewardDisable,IsRebirthSkillDisable,IsInvincibleAfterRevive off,StatusLimit=0
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.UpdateInstantStateFlagsWithMask"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@startBit",adUnsignedTinyInt,adParamInput,,8)
		Call .AppendParam("@bitLength",adUnsignedTinyInt,adParamInput,,8)
		Call .AppendParam("@value",adInteger,adParamInput,,0)
		blsResult = .ExecNoRecords()
	End with
	set sphn = Nothing
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_JoinSquare"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@squareName",adVarWChar,adParamInput,50,squareName)
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
