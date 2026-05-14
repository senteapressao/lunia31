<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init() 
	' request	: characterName
	' response	: @id
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim rs2
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterName,id
	If UBound(parameters)<0 Then Call Error("not enough parameter")	
	characterName	= Parameters(0)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.AddLottery"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@id",adBigInt,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sphn.frk_n4ErrorCode>0 Then
			Call Error(sphn.frk_n4ErrorCode)
		Else
			Call Error(sphn.frk_strErrorText)
		End If
	Else
		id = sphn.GetParamValue("@id")
	End If
	set sphn = Nothing

	Call Ok(id)
	
%>