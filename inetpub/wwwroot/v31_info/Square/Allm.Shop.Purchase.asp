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
	' request  : characterName,oidProduct
	' response : orderNumber

	Dim i,j,k
	Dim sph,RS,blsResult : blsResult = False

	' preparing request parameters
	Dim characterName,oidProduct,nOrderQuantity,oidCoupon_User
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	characterName	= Parameters(0)
	oidProduct		= Parameters(1)
	nOrderQuantity	= Parameters(2)
	oidCoupon_User	= Parameters(3)

	If oidCoupon_User="0" Then oidCoupon_User=null
	
	Dim accountName
	accountName = GetAccountName(characterName)
	'accountName = "indralig"

	If accountName="" OR IsEmpty(accountName) OR IsNull(accountName) Then Call Error(99)

	' Request To ShopDB for ProductInfo
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

	Dim oPublisher,Result
	Set oPublisher = new Publisher
	blsResult = False
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.public_Order_CheckConstraint_v2"
		Call .InitCommand()
		Call .AppendParam("@oidUser",adInteger,adParamInput,,oidUser)
		Call .AppendParam("@strUserID",adVarChar,adParamInput,50,strUserID)
		Call .AppendParam("@oidProduct",adInteger,adParamInput,,oidProduct)
		Call .AppendParam("@nOrderQuantity",adSmallint,adParamInput,,nOrderQuantity)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		'Println "success"
	Else
		If sph.frk_n4ErrorCode=200 Then
			Call Error(9)
		ElseIf sph.frk_n4ErrorCode=201 Then
			Call Error(4)
		ElseIf sph.frk_n4ErrorCode=202 Then
			Call Error(5)
		ElseIf sph.frk_n4ErrorCode=203 Then
			Call Error(6)
		ElseIf sph.frk_n4ErrorCode=204 Then
			Call Error(7)
		ElseIf sph.frk_n4ErrorCode=205 Then
			Call Error(8)
		End If
		Call Error(3)
	End If
	set sph = Nothing
	
	blsResult = False
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.public_Order_GetAmount_v2"
		Call .InitCommand()
		Call .AppendParam("@strUserID",adVarChar,adParamInput,20,strUserID) 'only TW
		Call .AppendParam("@oidProduct",adInteger,adParamInput,,oidProduct)
		Call .AppendParam("@nOrderQuantity",adSmallint,adParamInput,,nOrderQuantity)
		Call .AppendParam("@oidCoupon_User",adInteger,adParamInput,,oidCoupon_User)
		Call .AppendParam("@nCash",adInteger,adParamOutput,,null)
		Call .AppendParam("@nDiscount",adInteger,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		nTotalPrice = sph.GetParamValue("@nCash")
	Else
		Call Error(3)
	End If
	set sph = Nothing

	Dim TID
	TID = Left(FormatDt(Now, "ISO_TM") &"_"& strUserID &"_"& Left(GetGuid(),30),64)

	' Request To ShopDB for Purchase
	On Error Resume Next

	blsResult = False
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.public_Order_PurchaseProduct_v2"
		Call .InitCommand()
		Call .AppendParam("@oidUser",adInteger,adParamInput,,oidUser)
		Call .AppendParam("@strUserID",adVarChar,adParamInput,50,strUserID)
		Call .AppendParam("@TID",adChar,adParamInput,64,TID)
		Call .AppendParam("@nOrderMethod",adUnsignedTinyInt,adParamInput,,3)
		Call .AppendParam("@oidProduct",adInteger,adParamInput,,oidProduct)
		Call .AppendParam("@nOrderQuantity",adSmallint,adParamInput,,nOrderQuantity)
		Call .AppendParam("@oidCoupon_User",adInteger,adParamInput,,oidCoupon_User)
		Call .AppendParam("@oidOrder",adInteger,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If Err.Number<>0 Then blsResult=False
	On Error Goto 0

	If blsResult Then
		oidOrder = sph.GetParamValue("@oidOrder")
	Else
		Call Error(3)
	End If
	set sph = Nothing
	
	' use Bonus Point
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.public_Bonus_Use"
		Call .InitCommand()
		Call .AppendParam("@oidUser",adInteger,adParamInput,,oidUser)
		Call .AppendParam("@strUserID",adVarChar,adParamInput,50,strUserID)
		Call .AppendParam("@oidOrder",adInteger,adParamInput,,oidOrder)
		Call .AppendParam("@nPoint_Used",adInteger,adParamOutput,,null)
		Call .AppendParam("@nPoint_Remain",adInteger,adParamOutput,,null)
		Call .AppendParam("@nTotalPrice",adInteger,adParamOutput,,null)		
		blsResult = .ExecNoRecords()
	End with
	
	If blsResult Then
		Dim nPoint_Used,nPoint_Remain
		nPoint_Used = sph.GetParamValue("@nPoint_Used")
		nPoint_Remain = sph.GetParamValue("@nPoint_Remain")		
		nTotalPrice = sph.GetParamValue("@nTotalPrice")
	End If
	set sph = Nothing

	If nTotalPrice=0 Then
		Call Ok(characterName & SEPARATOR & accountName & SEPARATOR & oidOrder)
	End If

	' Billing
	Dim oidPayment,sql
	Set oPublisher = new Publisher
	oPublisher.DEBUG = False
	If oPublisher.use_money(strUserID,nTotalPrice,TID,oidProduct,nOrderQuantity,strProductName,nSalePrice,"buy","") Then
		oidPayment = oPublisher.strReturn
	Else
		' Rollback
		Set sph = new SPHelper
		with sph
		.DEBUG = False
			Set .cmd = Command
			.ConnStr = cashDBconnectionString
			.SPName = "dbo.public_Order_Cancel"
			Call .InitCommand()
			Call .AppendParam("@oidOrder",adInteger,adParamInput,,oidOrder)
			Call .AppendParam("@strReason",adVarchar,adParamInput,100,"billing fail")
			blsResult = .ExecNoRecords()
		End with
		set sph = Nothing
		
'		Set oPublisher = new Publisher
'		'oPublisher.DEBUG = True
'		Call oPublisher.rollback_money(strUserID,oidPayment,TID,nTotalPrice)
'		Set oPublisher = Nothing
		
		Set sph = new SPHelper
		with sph
		.DEBUG = False
			Set .cmd = Command
			.ConnStr = cashDBconnectionString
			.SPName = "dbo.public_Bonus_Cancel"
			Call .InitCommand()
			Call .AppendParam("@oidOrder",adInteger,adParamInput,,oidOrder)
			blsResult = .ExecNoRecords()
		End with
		set sph = Nothing

		sql = "insert dbo.tblOrder_Fail(logDate,oidOrder,strUserID,strReturn) select getdate(),?,?,?"
		with Command
			.ActiveConnection = cashDBconnectionString
			.CommandText = sql
			.CommandType = adCmdText
			.Parameters.Append .CreateParameter("@oidOrder",adInteger,adParamInput,,oidOrder)
			.Parameters.Append .CreateParameter("@strUserID",adVarWChar,adParamInput,50,strUserID)
			.Parameters.Append .CreateParameter("@strReturn",adVarWChar,adParamInput,300,oPublisher.strReturn)
			.Execute ,,adExecuteNoRecords
			Do Until Command.Parameters.Count = 0
				Command.Parameters.Delete (0)
			Loop
		End with

		Call Error(2)
	End If

	Call Ok(characterName & SEPARATOR & accountName & SEPARATOR & oidOrder)
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