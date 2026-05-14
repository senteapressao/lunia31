<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init() 
	' request	: characterSerial,achievementHash,score,hasReward
	' response	: achievementHash,hasReward
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	
	' preparing request parameters
	Dim characterSerial,achievementHash,score,hasReward
	If UBound(parameters)<3 Then Call Error("not enough parameter")	
	characterSerial		= Parameters(0)
	achievementHash		= Parameters(1)
	score				= Parameters(2)
	hasReward			= Parameters(3)

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.achievement_CompleteUnique"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE"	,adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterSerial",adBigInt,adParamInput,,characterSerial)
		Call .AppendParam("@achievementHash",adBigInt,adParamInput,,achievementHash)
		Call .AppendParam("@score"			,adInteger,adParamInput,,score)
		Call .AppendParam("@hasReward"		,adInteger,adParamInput,,hasReward)
		blsResult = .ExecNoRecords()
	End with

	Dim ret
	If blsResult Then
		ret = sphn.GetParamValue("RETURN_VALUE")
	End If
	set sphn = Nothing
	
	Call Ok(ret & SEPARATOR & achievementHash & SEPARATOR & hasReward)
%>