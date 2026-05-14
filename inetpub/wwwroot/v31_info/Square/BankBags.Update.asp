<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,bagNumber,expireDate
	' response : characterName,bagNumber,expireDate

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName,bagNumber,expireDate
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	bagNumber = Parameters(1)
	expireDate = Parameters(2)

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.BankBags_Update"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@bagNumber",adUnsignedTinyInt,adParamInput,,bagNumber)
		Call .AppendParam("@expireDate",adDBTimeStamp,adParamInput,,FormatDt(expireDate,"SQL_TM"))
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

	retString = characterName & SEPARATOR & bagNumber & SEPARATOR & FormatDt(expireDate,"SQL_TM")

	Call OK(retString)
%>