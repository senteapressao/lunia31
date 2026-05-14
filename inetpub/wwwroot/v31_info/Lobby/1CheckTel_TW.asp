<!--#include file="../common.asp"-->
<!--#include file="../soap_connect.asp"-->
<%
	' request  : account name
	' response : account name
	Call Init()

	' preparing request parameters
	Dim accountName
	
	accountName	= Parameters(0)

	Dim strSoapWSDL,strSoapAction,strSoapMethod,strSoapNS
	Dim soap,xmlDom
	Dim nodeResult

	strSoapWSDL = LoginWSDL
	strSoapAction = "http://tempuri.org/check_tel"
	strSoapMethod = "check_tel" 
	strSoapNS = "http://tempuri.org/"
	
	Set soap = new SoapConnect
	Set xmlDom = Server.CreateObject("Msxml2.DOMDocument.3.0")
	
	With soap
		.Debug = False
		.WSDL = strSoapWSDL
		.Action = strSoapAction
		.Method = strSoapMethod
		.NS = strSoapNS
		Call .AppendSoapXmlByValue("gameid",accountName)
		.Connect
	End With
	
	If Not(soap.objReader.Fault Is Nothing) Then
		Call Error("error occurs during soap connect")
	End If

	'Response.Write Replace(Replace(soap.objReader.DOM.xml,"<","&lt;"),">","&gt;") &"<br />"
	
	Set xmlDom = CreateObject("Msxml2.DOMDocument.3.0")
	xmlDom.async = False
	xmlDom.validateOnparse = False
	xmlDom.load(soap.objReader.DOM)
	
	Set nodeResult = xmlDom.selectSingleNode("//"& strSoapMethod &"Result")
	If nodeResult Is Nothing Then
		Call Error("return value is not valid")
	End If

	If nodeResult.Text="1" Then
		Call Error(1)
	ElseIf nodeResult.Text="2" Then
		Call Error(2)
	ElseIf nodeResult.Text="3" Then
		Call Error(3)
	ElseIf nodeResult.Text="0" Then
		Call Ok(accountName)
	End If
%>
