<%	Option Explicit %>
<!--#include virtual="./include/connect.asp"-->
<!--#include virtual="include/soap_connect.asp"-->
<!--#include file="./common.asp"-->
<%
	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim retString : retString = ""	
	Dim accountName,oidProduct
	Dim isVIP,point
	
	If Ubound(Parameters)<1 Then Call Error("not enough parameter")
	
	accountName	= Parameters(0)
	oidProduct	= Parameters(1)
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = ConnStrShop
		.SPName = "dbo.public_VIP_Check"
		Call .InitCommand()
		Call .AppendParam("@strUserID",adVarChar,adParamInput,20,accountName)
		Call .AppendParam("@isVIP",adUnsignedTinyInt,adParamInputOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		isVIP = sph.GetParamValue("@isVIP")
	Else
		isVIP = 0
	End If
	set sph = Nothing
	
	' Billing
	Dim strSoapWSDL,strSoapAction,strSoapMethod,strSoapNS
	Dim soap,xmlDom
	Dim nodeResult,arrResult
	
	strSoapWSDL = LoginWSDL
	strSoapAction = "http://tempuri.org/check_id_point"
	strSoapMethod = "check_id_point" 
	strSoapNS = "http://tempuri.org/"
	
	Set soap = new SoapConnect
	Set xmlDom = Server.CreateObject("Msxml2.DOMDocument.3.0")
	
	With soap
		.DEBUG = False
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
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = ConnStrShop
		.SPName = "dbo.public_Product_Preview_GetInfo"
		Call .InitCommand()
		Call .AppendParam("@oidProduct",adInteger,adParamInput,,oidProduct)
		blsResult = .ExecRecordset2()
	End with

	If blsResult Then
		Set RS = sph.rs
		Set RS2 = sph.rs2
	Else
		Call Error("data error")
	End If
	set sph = Nothing
	
	If Not(RS.Eof) Then
	
		If RS2.Eof Then
			retString = retString & SEPARATOR & RS("oidProduct") &_
					SUBSEPARATOR & RS("strProductCode") &_
					SUBSEPARATOR & RS("strProductName") &_
					SUBSEPARATOR & RS("nProductExpire") &_
					SUBSEPARATOR & RS("nProductEA") &_
					SUBSEPARATOR & RS("nSalePrice") &_
					SUBSEPARATOR & RS("nSalePriceVIP")
		Else
			Do Until RS2.Eof
				retString = retString & SEPARATOR & RS2("oidProduct") &_
					SUBSEPARATOR & RS2("strProductCode") &_
					SUBSEPARATOR & RS2("strProductName") &_
					SUBSEPARATOR & RS2("nProductExpire") &_
					SUBSEPARATOR & RS2("nProductEA") &_
					SUBSEPARATOR & RS2("nSalePrice") &_
					SUBSEPARATOR & RS2("nSalePriceVIP")

				RS2.MoveNext : i=i+1
			Loop
			RS2.Close			
		End If
	End If
	RS.Close

	Call Ok(retString)
%>