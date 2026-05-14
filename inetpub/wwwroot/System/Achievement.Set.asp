<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init() 
	' request	: characterName,classNumber,achievementPoint,isComplete(1:complete,0:working),characterSerial,achievementHash,(logDate)
	' response	: nothing
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	
	' preparing request parameters
	Dim characterName,classNumber,achievementPoint,isComplete,characterSerial,achievementHash,logDate
	If UBound(parameters)<5 Then Call Error("not enough parameter")	
	characterName		= Parameters(0)
	classNumber			= Parameters(1)
	achievementPoint	= Parameters(2)
	isComplete			= Parameters(3)
	characterSerial		= Parameters(4)
	achievementHash		= Parameters(5)
	If isComplete = "1" Then
		logDate			= Parameters(6)
		logDate	= FormatDt(logDate,"SQL_TM")
	End If
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.achievementSet"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@classNumber",adSmallInt,adParamInput,,classNumber)
		Call .AppendParam("@achievementPoint",adInteger,adParamInput,,achievementPoint)
		Call .AppendParam("@isComplete",adTinyInt,adParamInput,,isComplete)
		Call .AppendParam("@characterSerial",adBigInt,adParamInput,,characterSerial)
		Call .AppendParam("@achievementHash",adBigInt,adParamInput,,achievementHash)
		If isComplete = "1" Then
			Call .AppendParam("@logDate",adDBTimeStamp,adParamInput,20,logDate)
		Else
			Call .AppendParam("@logDate",adDBTimeStamp,adParamInput,20,null)
		End If
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
	
	Call Ok(0)
%>