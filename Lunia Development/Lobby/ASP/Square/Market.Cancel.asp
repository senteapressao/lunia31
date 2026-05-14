<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,marketId
	' response : characterName,marketId,itemHash,stackedCount,itemSerial,instance,itemExpire,salePrice,expireDate

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName,marketId
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	characterName	= Parameters(0)
	marketId		= Parameters(1)

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.Market_Cancel"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@marketId",adBigInt,adParamInput,,marketId)
		blsResult = .ExecRecordset()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		Set rs = sphn.rs
	Else
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
	End If
	set sphn = Nothing

	If Not(rs.Eof) Then
		retString = retString & rs(0) &_
			SEPARATOR & rs(1) & SEPARATOR & rs(2) &_
			SEPARATOR & rs(3) & SEPARATOR & rs(4) &_
			SEPARATOR & FormatDt(rs(5),"SQL_TM") &_
			SEPARATOR & rs(6) &_
			SEPARATOR & FormatDt(rs(7),"SQL_TM")
	End If
	rs.close

	retString = characterName & SEPARATOR & retString

	Call OK(retString)
%>