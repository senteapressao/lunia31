<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName(buyer),marketId
	' response : characterName(buyer)
	'			,marketId,characterName,itemType,itemHash,stackedCount,itemSerial,instance,itemExpire
	'			,itemName,petName,petLevel,isRare,rareProbability,enchantSerial
	'			,salePrice,expireDate,regDate

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
		.SPName = "dbo.Market_Contract"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
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
		retString = retString & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) &_
			SEPARATOR & rs(3) & SEPARATOR & rs(4) & SEPARATOR & rs(5) &_
			SEPARATOR & rs(6) &_
			SEPARATOR & FormatDt(rs(7),"SQL_TM") &_
			SEPARATOR & rs(8) &_
			SEPARATOR & rs(9) & SEPARATOR & rs(10) & SEPARATOR & rs(11) &_
			SEPARATOR & rs(12) & SEPARATOR & rs(13) & SEPARATOR & rs(14) &_
			SEPARATOR & FormatDt(rs(15),"SQL_TM") & SEPARATOR & FormatDt(rs(16),"SQL_TM") &_
			SEPARATOR & rs(17)
	End If
	rs.close

	retString = characterName & SEPARATOR & retString

	Call OK(retString)
%>