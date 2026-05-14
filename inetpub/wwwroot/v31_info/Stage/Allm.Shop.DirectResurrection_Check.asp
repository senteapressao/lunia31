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
	' response : characterName

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
	
	Dim nCash

	' Billing
	Dim oPublisher,Result
	Set oPublisher = new Publisher
	If oPublisher.get_money(strUserID) Then
		nCash = oPublisher.strReturn
	Else
		'Call ThrowError(null,oPublisher.strErrorText,"go:/default.asp")
		nCash = 0
	End If
	Set oPublisher = Nothing
	
	'Z Point
	Dim nPoint
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.public_Bonus_GetInfo"
		Call .InitCommand()
		Call .AppendParam("@oidUser",adInteger,adParamInput,,0) 'only KR
		'Call .AppendParam("@oidUser",adInteger,adParamInput,,oidUser)
		Call .AppendParam("@strUserID",adVarChar,adParamInput,50,strUserID)
		Call .AppendParam("@nPoint",adInteger,adParamInputOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		nPoint = sph.GetParamValue("@nPoint")
	Else
		nPoint = 0
	End If
	set sph = Nothing
	
	If CLng(nCash)+CLng(nPoint) < CLng(nTotalPrice) Then
		Call Error(2)
	End If 

	Call Ok(characterName & SEPARATOR & accountName & SEPARATOR & CLng(nCash)+CLng(nPoint))
%>