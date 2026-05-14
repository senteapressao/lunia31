<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init() 
	' request	: characterName,amount
	' response	: Ok
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim rs2
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterName, amount, coin_type
	If UBound(parameters)<2 Then Call Error("not enough parameter")	
	characterName	= Parameters(0)
	amount	= Parameters(1)
	coin_type	= Parameters(2)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.AddLcoin"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@amount",adInteger,adParamInput,,amount)
		Call .AppendParam("@type",adInteger,adParamInput,,coin_type)
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
	
	Call Ok(characterName & SEPARATOR & amount & SEPARATOR & coin_type)
	
%>