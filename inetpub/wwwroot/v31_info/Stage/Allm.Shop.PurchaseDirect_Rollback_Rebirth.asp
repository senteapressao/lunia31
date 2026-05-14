<!--#include file="../common.asp"-->
<!--#include file="../soap_connect.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Server.Scripttimeout = 15

	Call Init()
	' request  : orderNumber
	' response : orderNumber

	Dim i,j,k
	Dim sph,RS,blsResult : blsResult = False
	
	' preparing request parameters
	Dim orderNumber
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	orderNumber = Parameters(0)
	
	If orderNumber="" Then Call Error("no orderNumber")

	' Request To ShopDB for Rollback Purchase
	Dim oidOrder,strReason,oidPayment,strUserID,nTotalPrice
	oidOrder	= orderNumber
	oidPayment	= 0
	strReason	= "rollback by game error"
    
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.public_Order_PurchaseDirect_Rollback"
		Call .InitCommand()
		Call .AppendParam("@oidOrder",adInteger,adParamInput,,oidOrder)
		Call .AppendParam("@strReason",adVarWChar,adParamInput,100,strReason)
		Call .AppendParam("@oidPayment",adVarChar,adParamInputOutput,50,null)
		Call .AppendParam("@strUserID",adVarChar,adParamInputOutput,50,null)
		Call .AppendParam("@nTotalPrice",adInteger,adParamInputOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		oidPayment = sph.GetParamValue("@oidPayment")
		strUserID = sph.GetParamValue("@strUserID")
		nTotalPrice = sph.GetParamValue("@nTotalPrice")
	Else
		oidPayment = sph.GetParamValue("@oidPayment")
		Call Ok( 1 & SEPARATOR & oidOrder & SEPARATOR & oidPayment )
	End If
	set sph = Nothing

	' Rollback Billing
	Dim strSoapWSDL,strSoapAction,strSoapMethod,strSoapNS
	Dim soap,xmlDom
	Dim nodeResult,arrResult,strResult

	blsResult  = False
	On Error Resume Next

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
		Call .AppendSoapXmlByValue("gameid",strUserID)
		Call .AppendSoapXmlByValue("addpoint",nTotalPrice)
		Call .AppendSoapXmlByValue("unicodeno",oidPayment)
		.Connect
	End With
	
	If Not(soap.objReader.Fault Is Nothing) Then
		Call RollBack_Fail(oidOrder)
	End If

	Set xmlDom = CreateObject("Msxml2.DOMDocument.3.0")
	xmlDom.async = False
	xmlDom.validateOnparse = False
	xmlDom.load(soap.objReader.DOM)

	Set nodeResult = xmlDom.selectSingleNode("//"& strSoapMethod &"Result")
	If nodeResult Is Nothing Then
		Call RollBack_Fail(oidOrder)
	End If

	If Left(nodeResult.Text,1)<>"0" Then
		Call RollBack_Fail(oidOrder)
	End If

	If Err.Number<>"0" Then Call RollBack_Fail(oidOrder)
	On Error Goto 0

	Call Ok( 0 & SEPARATOR & oidOrder & SEPARATOR & oidPayment )
%>
<%
	Sub RollBack_Fail(oidOrder)
		Set sph = new SPHelper
		with sph
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = cashDBconnectionString
			.SPName = "dbo.public_Order_PurchaseDirect_Rollback_Fail"
			Call .InitCommand()
			Call .AppendParam("@oidOrder",adInteger,adParamInput,,oidOrder)
			blsResult = .ExecNoRecords()
		End with
		set sph = Nothing
		Call Ok( 2 & SEPARATOR & oidOrder & SEPARATOR & oidPayment )
	End Sub
%>