<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()	
	' request  : accountName
	' response : keyString
	
	Dim sphn,blsResult : blsResult = False
	Dim retString : retString = ""

	Dim accountName
	accountName	= Parameters(0)
	Dim keyString
	
	' call Stored procedure	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.GetKeySetting"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		Call .AppendParam("@keyString",adVarWChar,adParamOutput,3000,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If		
		keyString = sphn.GetParamValue("@keyString")
	End If
	set sphn = Nothing
	
	retstring = keyString
	Call Ok(retString)
%>