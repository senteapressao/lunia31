<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init() 
	' request	: characterName,license
	' response	: Ok
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim rs2
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterName, license
	If UBound(parameters)<1 Then Call Error("not enough parameter")	
	characterName	= Parameters(0)
	license	= Parameters(1)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.AddItemEffectLicense"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@license",adVarWChar,adParamInput,50,license)
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
	
	Call Ok(characterName & SEPARATOR & license)
	
%>