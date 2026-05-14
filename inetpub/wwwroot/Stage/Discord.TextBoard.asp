<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	Dim httpRequest, postResponse, data
	data	= Parameters(0)
	Set httpRequest = Server.CreateObject("MSXML2.ServerXMLHTTP.3.0")
	httpRequest.Open "POST", "https://discord.com/api/webhooks/1048003730500497408/jGHyuDASa0bTHrs4gi2FIMCznJpKf30l67FK_ORUB70dJtSdsV_MGWAH2OAPE7Zt85NU", False
	httpRequest.SetRequestHeader "Content-Type", "application/json"

	httpRequest.Send data

	postResponse = httpRequest.ResponseText

	Response.Write postResponse ' or do something else with it
	
%>