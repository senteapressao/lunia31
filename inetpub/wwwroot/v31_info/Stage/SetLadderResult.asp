<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,classNumber,ladderPoint,isWin,subInfo
	' response : nothing

	Dim i,j,k
	Dim sphn,blsResult : blsResult = False
	Dim retString : retString = ""
	
	' preparing request parameters	
	Dim characterName,classNumber,ladderPoint,isWin,subInfo
	If UBound(parameters)<3 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	classNumber = Parameters(1)
	ladderPoint = Parameters(2)
	isWin = Parameters(3)
	subInfo = Parameters(4)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.LadderUpdateResult"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@classNumber",adSmallInt,adParamInput,,classNumber)
		Call .AppendParam("@ladderPoint",adInteger,adParamInput,,ladderPoint)
		Call .AppendParam("@isWin",adTinyInt,adParamInput,,isWin)
		Call .AppendParam("@subInfo",adTinyInt,adParamInput,,subInfo)
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

	Call Ok("Ok")
%>