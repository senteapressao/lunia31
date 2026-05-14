<!--#include file="../common.asp"-->
<!--#include file="../soap_connect.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Server.ScriptTimeOut = 15

	Call Init()
	' request  : characterName
	' response : orderNumber

	Dim i,j,k
	Dim sph,RS,blsResult : blsResult = False

	' preparing request parameters
	Dim characterName
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	
	Dim accountName
	accountName=GetAccountName(characterName)
	'accountName="luniaob014"

	If accountName="" Or IsEmpty(accountName) Or IsNull(accountName) Then Call Error("no accountName")

	' Request To ShopDB for ProductInfo
	Dim oidProduct,nOrderQuantity : oidProduct = 109 : nOrderQuantity = 1
	Dim strProductName,nSalePrice	
	Dim nTotalPrice
	Dim oidUser,strUserID,oidOrder
	oidUser	= 0
	strUserID = accountName

	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.public_Product_GetInfo"
		Call .InitCommand()
		Call .AppendParam("@oidProduct",adInteger,adParamInput,,oidProduct)
		blsResult = .ExecRecordset()
	End with

	If blsResult Then
		Set RS = sph.rs
		If RS.Eof Then Call Error(3)
		strProductName = RS("strProductName")
		nSalePrice = RS("nSalePrice")
	Else
		Call Error(1)
	End If
	set sph = Nothing

	nTotalPrice = nSalePrice * nOrderQuantity

	Dim TID
	TID = Left(FormatDt(Now, "ISO_TM") &"_"& strUserID &"_"& Left(GetGuid(),30),64)

	' Billing
	Dim oidPayment
	Dim strSoapWSDL,strSoapAction,strSoapMethod,strSoapNS
	Dim soap,xmlDom
	Dim nodeResult,arrResult,strResult

	strSoapWSDL = LoginWSDL
	strSoapAction = "http://tempuri.org/buy_lunia"
	strSoapMethod = "buy_lunia" 
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
		Call .AppendSoapXmlByValue("calpoint",nTotalPrice)
		Call .AppendSoapXmlByValue("itemno",oidProduct)
		Call .AppendSoapXmlByValue("itemsize",nOrderQuantity)
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

	If Left(nodeResult.Text,1)="0" Then
		arrResult = Split(nodeResult.Text,"|")
		If IsArray(arrResult)=False Then Error("error occurs")
		If Ubound(arrResult)<3 Then Error("error occurs")
		oidPayment = arrResult(3)
	Else
		Call Error(2)
	End If

	' Request To ShopDB for Purchase
	Dim strProductInfo : strProductInfo	= oidProduct &","& nOrderQuantity

	On Error Resume Next

	blsResult = False
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.public_Order_PurchaseDirect"
		Call .InitCommand()
		Call .AppendParam("@oidUser",adInteger,adParamInput,,oidUser)
		Call .AppendParam("@strUserID",adVarChar,adParamInput,20,strUserID)
		Call .AppendParam("@TID",adChar,adParamInput,64,TID)
		Call .AppendParam("@strProductInfo",adVarChar,adParamInput,1000,strProductInfo)
		Call .AppendParam("@oidPayment",adVarChar,adParamInput,50,oidPayment)
		Call .AppendParam("@oidOrder",adInteger,adParamInputOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If Err.Number<>0 Then blsResult=False
	On Error Goto 0

	If blsResult Then
		oidOrder = sph.GetParamValue("@oidOrder")
		Call Ok(oidOrder)
	Else
		' Rollback Billing
		strSoapWSDL = LoginWSDL
		strSoapAction = "http://tempuri.org/return_lunia"
		strSoapMethod = "return_lunia" 
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
			Call .AppendSoapXmlByValue("addpoint",nTotalPrice)
			Call .AppendSoapXmlByValue("unicodeno",oidPayment)
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

		Call Error(3)
	End If
	set sph = Nothing
%>
<%
	Function GetGuid()
		Dim objTypeLib, strGuid
		Set objTypeLib = Server.CreateObject("Scriptlet.Typelib")
			strGuid = objTypeLib.Guid
		Set objTypeLib = Nothing
		strGuid = Replace(strGuid,"-","")
		strGuid = Replace(strGuid, "{", "")
		strGuid = Replace(strGuid, "}", "")
		GetGuid = strGuid
		Exit Function
	End Function
%>
