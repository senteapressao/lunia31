<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,basic,pcRoom,extra,extraAmount,extraExpire,initDate
	' response : characterName

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName,basic,item,evt,pcRoom,extra,extraAmount,extraExpire,initDate
	If UBound(parameters)<8 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	basic = Parameters(1)
	item = Parameters(2)
	evt = Parameters(3)
	pcRoom = Parameters(4)
	extra = Parameters(5)
	extraAmount = Parameters(6)
	extraExpire = Parameters(7)
	initDate = Parameters(8)

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.Token_Update"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@basic",adInteger,adParamInput,,basic)
		Call .AppendParam("@item",adInteger,adParamInput,,item)
		Call .AppendParam("@event",adInteger,adParamInput,,evt)
		Call .AppendParam("@pcRoom",adInteger,adParamInput,,pcRoom)
		Call .AppendParam("@extra",adInteger,adParamInput,,extra)
		Call .AppendParam("@extraAmount",adInteger,adParamInput,,extraAmount)
		Call .AppendParam("@extraExpire",adDBTimeStamp,adParamInput,,extraExpire)
		Call .AppendParam("@initDate",adDBTimeStamp,adParamInput,,initDate)
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

	retString = characterName

	Call OK(retString)
%>