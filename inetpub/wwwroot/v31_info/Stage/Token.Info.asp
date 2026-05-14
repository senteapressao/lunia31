<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	' response : characterName,basic,item,evt,pcRoom,extra,extraAmount,extraExpire,initDate

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	characterName = Parameters(0)

	Dim basic,item,evt,pcRoom,extra,extraAmount,extraExpire,initDate

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.Token_Info"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@basic",adInteger,adParamOutput,,null)
		Call .AppendParam("@item",adInteger,adParamOutput,,null)
		Call .AppendParam("@event",adInteger,adParamOutput,,null)
		Call .AppendParam("@pcRoom",adInteger,adParamOutput,,null)
		Call .AppendParam("@extra",adInteger,adParamOutput,,null)
		Call .AppendParam("@extraAmount",adInteger,adParamOutput,,null)
		Call .AppendParam("@extraExpire",adDBTimeStamp,adParamOutput,,null)
		Call .AppendParam("@initDate",adDBTimeStamp,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		basic = sphn.GetParamValue("@basic")
		item = sphn.GetParamValue("@item")
		evt = sphn.GetParamValue("@event")
		pcRoom = sphn.GetParamValue("@pcRoom")
		extra = sphn.GetParamValue("@extra")
		extraAmount = sphn.GetParamValue("@extraAmount")
		extraExpire = sphn.GetParamValue("@extraExpire")
		initDate = sphn.GetParamValue("@initDate")
	End If
	set sphn = Nothing

	retString = characterName &_
			SEPARATOR & basic & SEPARATOR & item & SEPARATOR & evt & SEPARATOR & pcRoom &_
			 SEPARATOR & extra & SEPARATOR & extraAmount & SEPARATOR & FormatDt(extraExpire,"SQL_TM") &_
			SEPARATOR & FormatDt(initDate,"SQL_TM")

	Call OK(retString)
%>