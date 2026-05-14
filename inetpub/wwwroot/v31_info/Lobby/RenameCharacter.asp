<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : accountName
	' response : characterName_old,characterName_new....
		
	Dim sphn,blsResult : blsResult = False
	Dim retString : retString = ""

	Dim accountName
	accountName	= Parameters(0)
	Dim characterList
	
	' call Stored procedure	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.RenameCharacterLoop"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		Call .AppendParam("@characterList",adVarWChar,adParamOutput,1000,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If		
		characterList = sphn.GetParamValue("@characterList")
	End If
	set sphn = Nothing
	
	retstring = characterList
	retstring = replace(retstring,"|",SEPARATOR)
	retstring = replace(retstring,",",SUBSEPARATOR)
	Call Ok(retString)
%>