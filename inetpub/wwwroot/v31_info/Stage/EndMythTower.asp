<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,classNumber,stageLevel,floor,playTime
	' response : 
		
	Dim sphn,blsResult : blsResult = False
	Dim retString : retString = ""	

	Dim characterName,classNumber,stageLevel,floor,playTIme
	characterName	= Parameters(0)
	classNumber		= Parameters(1)
	stageLevel		= Parameters(2)
	floor			= Parameters(3)
	playTIme		= Parameters(4)
	
	'response 변수 초기화
	
	' call Stored procedure	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.InsertMythTowerPlayTime"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@classNumber",adSmallInt,adParamInput,,classNumber)
		Call .AppendParam("@stageLevel",adSmallInt,adParamInput,,stageLevel)
		Call .AppendParam("@floor",adTinyInt,adParamInput,,floor)
		Call .AppendParam("@playTiIme",adBigInt,adParamInput,,playTime)

		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
	End If
	set sphn = Nothing
	
	Call Ok(characterName)
%>