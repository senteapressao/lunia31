<!--#include file="../common.asp"-->
<!--#include file="../md5_class.asp"-->
<!--#include file="../publisher.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Server.ScriptTimeOut = 15
	
	Call Init()
	' request  : characterName
	' response : characterName

	Dim i,j,k
	Dim sphn
	Dim sph,RS,blsResult : blsResult = False

	' preparing request parameters
	Dim characterName
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	characterName = Parameters(0)

	Dim freeResCount 
	
	Dim accountName
	accountName = GetAccountName(characterName)

	'STEP 1 : GET NUMBER OF REZZES AVAILABLE
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.GetResurrections"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		Call .AppendParam("@resultCode",adInteger,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		freeResCount = sphn.GetParamValue("@resultCode")
	End If
	set sphn = Nothing
	
	
	IF freeResCount = "-99" Then
		Call Error(2)
	Else
		Call Ok(characterName & SEPARATOR & accountName & SEPARATOR & "0")
	End If
	
	
	
	
%>