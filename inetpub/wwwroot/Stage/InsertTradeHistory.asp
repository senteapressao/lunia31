<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request	: gold,sender,[Hash,StackedCount,Serial,Instance,ExpireDate],target,[Hash,StackedCount,Serial,Instance,ExpireDate]
	' response	: 0
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim rs2
	
	' preparing request parameters
	Dim gold,sender,target,itemsSender,itemsTarget
	If UBound(parameters)<4 Then Call Error("not enough parameter")	
	gold		= Parameters(0)
	sender		= Parameters(1)
	itemsSender	= Parameters(2)
	target		= Parameters(3)
	itemsTarget	= Parameters(4)
	
	Dim lastMail,lastTrade
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.InsertTradeHistory"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@gold",adInteger,adParamInput,,gold)
		Call .AppendParam("@sender",adVarWChar,adParamInput,50,sender)
		Call .AppendParam("@itemsSender",adVarWChar,adParamInput,2000,itemsSender)
		Call .AppendParam("@target",adVarWChar,adParamInput,50,target)
		Call .AppendParam("@itemsTarget",adVarWChar,adParamInput,2000,itemsTarget)
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


	Call OK()
%>