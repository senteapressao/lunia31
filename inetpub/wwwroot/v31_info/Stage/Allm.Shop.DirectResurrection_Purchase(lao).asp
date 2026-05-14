<!--#include file="../common.asp"-->
<!--#include file="../md5_class.asp"-->
<!--#include file="../publisher.asp"-->
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
	accountName = GetAccountName(characterName)
	'accountName = "indralig"

	If accountName="" OR IsEmpty(accountName) OR IsNull(accountName) Then Call Error("no accountName")

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
		If RS.Eof Then Call Error(1)
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
	Dim oPublisher
	Dim oidPayment
	Set oPublisher = new Publisher
	'oPublisher.DEBUG = True
	If oPublisher.use_money(strUserID,nTotalPrice,TID,oidProduct,nOrderQuantity,strProductName,nSalePrice,"buy","") Then
		oidPayment = oPublisher.strReturn
	Else
		Call Error(2)
	End If
	Set oPublisher = Nothing

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
		Call .AppendParam("@strUserID",adVarChar,adParamInput,50,strUserID)
		Call .AppendParam("@TID",adChar,adParamInput,64,TID)
		Call .AppendParam("@strProductInfo",adVarChar,adParamInput,1000,strProductInfo)
		Call .AppendParam("@oidPayment",adVarChar,adParamInput,50,0)
		Call .AppendParam("@oidOrder",adInteger,adParamInputOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If Err.Number<>0 Then blsResult=False
	On Error Goto 0

	If blsResult Then
		oidOrder = sph.GetParamValue("@oidOrder")
		Call Ok(characterName & SEPARATOR & accountName & SEPARATOR & oidOrder)
	Else
		Set oPublisher = new Publisher
		'oPublisher.DEBUG = True
		Call oPublisher.rollback_money(strUserID,oidPayment,TID,nTotalPrice)
		Set oPublisher = Nothing
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
