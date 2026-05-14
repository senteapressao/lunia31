<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : mailId,characterName,flag
	' response : 0,mailId,flag
	
	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	
	' preparing request parameters
	Dim mailId,characterName,flag
	If UBound(Parameters)<2 Then Call Error("not enough parameter")
	mailId = Parameters(0)
	characterName = Parameters(1)
	flag = Parameters(2)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.PopMailItems"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@mailId",adInteger,adParamInput,,mailId)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@flag",adUnsignedTinyInt,adParamInput,,flag)
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
	
	Call Ok("0"& SEPARATOR & mailId & SEPARATOR & flag )
%>