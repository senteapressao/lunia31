<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,isBank,bagNumber,expireDate
	' response :

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim characterName,isBank,bagNumber,expireDate
	
	If UBound(Parameters)<3 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	isBank = Parameters(1)
	bagNumber = Parameters(2)
	expireDate = Parameters(3)

	' call Stored procedure
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.UpdateBagState"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@isBank",adBoolean,adParamInput,,isBank)
		Call .AppendParam("@bagNumber",adInteger,adParamInput,,bagNumber)		
		Call .AppendParam("@expireDate",adDBTimeStamp,adParamInput,,FormatDt(expireDate,"SQL_TM"))
		blsResult = .ExecNoRecords()
	End with

	Dim ret
	ret = sphn.GetParamValue("RETURN_VALUE")
	if ret<>0 Then
		Call Error(ret)
	End If
	set sphn = Nothing

	Call Ok(ret)
%>