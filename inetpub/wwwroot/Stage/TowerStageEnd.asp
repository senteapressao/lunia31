<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	'Parameters = Split("Moki1385191410020693600","")
	Call Init()
	' request  : characterName,classNumber,stageLevel,stageGroupHash,accessLevel,floor,playTime
	' response : 
		
	Dim sphn,blsResult : blsResult = False
	Dim retString : retString = ""	

	
	
	Dim characterName,classNumber,stageLevel,stageGroupHash,accessLevel,floor,playTIme
	If UBound(parameters)<6 Then Call Error("not enough parameter")
	characterName	= Parameters(0)
	classNumber		= Parameters(1)
	stageLevel		= Parameters(2)
	stageGroupHash	= Parameters(3)
	accessLevel		= Parameters(4)
	floor			= Parameters(5)
	playTIme		= Parameters(6)
	
	' call Stored procedure	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = True
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.TowerInsertPlayTime"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@classNumber",adSmallInt,adParamInput,,classNumber)
		Call .AppendParam("@stageLevel",adSmallInt,adParamInput,,stageLevel)
		Call .AppendParam("@stageGroupHash",adInteger,adParamInput,,stageGroupHash)
		Call .AppendParam("@accessLevel",adSmallInt,adParamInput,,accessLevel)
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