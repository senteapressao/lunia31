<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,targetName
	' response : 
	
	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterName,targetName
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	targetName = 	Parameters(1)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.SPName = "dbo.AddSpectate"
		.ConnStr = characterDBconnectionString
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@targetName",adVarWChar,adParamInput,50,targetName)
		blsResult = .ExecRecordset()
	End with
	
	set sphn = Nothing
	
	Call OK("0")
%>