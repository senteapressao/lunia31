<!--#include file="./common.asp"-->
<!--#include file="./soap_connect.asp"-->
<!--#include file="./DBhelper.asp"-->
<!--#include file="./DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim retString : retString = ""	
	Dim accountName,oidProduct
	Dim isVIP,point
	
	Call Init()
	If Ubound(Parameters)<1 Then Call Error("not enough parameter")
	
	accountName	= Parameters(0)
	oidProduct	= Parameters(1)
	
	' Billing
	Dim strSoapWSDL,strSoapAction,strSoapMethod,strSoapNS
	Dim soap,xmlDom
	Dim nodeResult,arrResult
	
	strSoapWSDL = LoginWSDL
	strSoapAction = "http://tempuri.org/check_id_VIP"
	strSoapMethod = "check_id_VIP" 
	strSoapNS = "http://tempuri.org/"
	
	Set soap = new SoapConnect
	Set xmlDom = Server.CreateObject("Msxml2.DOMDocument.3.0")
	
	With soap
		.DEBUG = True
		.WSDL = strSoapWSDL
		.Action = strSoapAction
		.Method = strSoapMethod
		.NS = strSoapNS
		Call .AppendSoapXmlByValue("gamename",accountName)
		.Connect
	End With
	
	If Not(soap.objReader.Fault Is Nothing) Then
		Call Error("error occurs during soap connect")
	End If

	Set xmlDom = CreateObject("Msxml2.DOMDocument.3.0")
	xmlDom.async = False
	xmlDom.validateOnparse = False
	xmlDom.load(soap.objReader.DOM)
	
	Set nodeResult = xmlDom.selectSingleNode("//"& strSoapMethod &"Result")
	If nodeResult Is Nothing Then
		Call Error("return value is not valid")
	End If
	
	isVIP = nodeResult.Text

	strSoapWSDL = LoginWSDL
	strSoapAction = "http://tempuri.org/check_id_point"
	strSoapMethod = "check_id_point" 
	strSoapNS = "http://tempuri.org/"
	
	Set soap = new SoapConnect
	Set xmlDom = Server.CreateObject("Msxml2.DOMDocument.3.0")
	
	With soap
		.DEBUG = True
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

	Set xmlDom = CreateObject("Msxml2.DOMDocument.3.0")
	xmlDom.async = False
	xmlDom.validateOnparse = False
	xmlDom.load(soap.objReader.DOM)
	
	Set nodeResult = xmlDom.selectSingleNode("//"& strSoapMethod &"Result")
	If nodeResult Is Nothing Then
		Call Error("return value is not valid")
	End If
	
	point = nodeResult.Text

	retString = retString & SEPARATOR & isVIP
	retString = retString & SEPARATOR & point

	Call Ok(retString)
%>