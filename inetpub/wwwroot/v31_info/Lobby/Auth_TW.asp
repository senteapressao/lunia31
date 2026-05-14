<!--#include file="../common.asp"-->
<!--#include file="../soap_connect.asp"-->
<!--#include file="../md5.asp"-->
<%
	' request  : account name, password, ip address
	' response : account name
	Call Init()
	If UBound(parameters)<2 Then Call Error("not enough parameter") 

	' preparing request parameters
	Dim accountName, password, ip
	
	accountName	= Parameters(0)
	password	= Parameters(1)
	ip			= Parameters(2)

'	params=Array(ServerGroupCode)
'	ret=ExecSP(gateDBconnectionString, "GetConnections", params)

'	If params(0)>10000 Then
'		'server connecions over 10000
'		Call Error(9)
'	End If
	
	Dim strSoapWSDL,strSoapAction,strSoapMethod,strSoapNS
	Dim soap,xmlDom
	Dim nodeResult

	strSoapWSDL = LoginWSDL
	strSoapAction = "http://tempuri.org/check_login"
	strSoapMethod = "check_login" 
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
		Call .AppendSoapXmlByValue("gamepass",password)
		Call .AppendSoapXmlByValue("ip",ip)
		.Connect
	End With
	
	If Not(soap.objReader.Fault Is Nothing) Then
		Call Error("error occurs during soap connect")
	End If
	
'	Response.Write Replace(Replace(soap.objReader.DOM.xml,"<","&lt;"),">","&gt;") &"<br />"
	
	Set xmlDom = CreateObject("Msxml2.DOMDocument.3.0")
	xmlDom.async = False
	xmlDom.validateOnparse = False
	xmlDom.load(soap.objReader.DOM)
	
	Set nodeResult = xmlDom.selectSingleNode("//"& strSoapMethod &"Result")
	If nodeResult Is Nothing Then
		Call Error("return value is not valid")
	End If

	If nodeResult.Text="-1" Then
		'not input accountName
		Call Error(11)
	ElseIf nodeResult.Text="-2" Then
		'not input password
		Call Error(12)
	ElseIf nodeResult.Text="-3" Then
		'not input ip
		Call Error(13)
	ElseIf nodeResult.Text="-4" Then
		'accountName is not 10byte
		Call Error(14)
	ElseIf nodeResult.Text="-5" Then
		'password is not 32byte
		Call Error(15)
	ElseIf nodeResult.Text="-6" Then
		'ip is not 15bytes
		Call Error(16)
	ElseIf nodeResult.Text="-8" Then
		'accountName left 2byte is not valid
		Call Error(18)
	ElseIf nodeResult.Text="-9" Then
		'error occurs during search account's password
		Call Error(19)
	ElseIf nodeResult.Text="0" Then
		'accountName or password is not valid
		Call Error(10)
	ElseIf nodeResult.Text="1" Then
		'success
	Else
		'etc error occurs
		Call Error(10)
	End If

	' call Stored procedure	
	Dim params, ret
	params=Array(REQUESTERID, accountName)
	
	ret=ExecSP(characterDBconnectionString, "NexonAuth", params)
	If ret=0 Then
		Call Ok(accountName)
	Else
		Call Error(ret)
	End If
%>
