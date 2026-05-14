<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : character name, class number
	' response : completeCount
	
	Dim sphn,blsResult : blsResult = False
	Dim retString : retString = ""	
	
	' preparing request parameters
	Dim characterName,questHash
	Dim completeCount
	If UBound(parameters)<1 Then Call Error("not enough parameter")	
	characterName	= Parameters(0)
	questHash		= Parameters(1)
	
	completeCount	= 0
	
	' call Stored procedure	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.GetCompletedQuestCount"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@questHash",adInteger,adParamInput,,questHash)
		Call .AppendParam("@completeCount",adInteger,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with
	
	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		completeCount = sphn.GetParamValue("@completeCount")
	End If
	set sphn = Nothing
	
	Call Ok(questHash & SEPARATOR & completeCount)
%>