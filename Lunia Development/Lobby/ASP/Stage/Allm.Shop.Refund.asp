<!--#include file="../common.asp"-->
<!--#include file="../md5_class.asp"-->
<!--#include file="../publisher.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Server.ScriptTimeOut = 15

	Call Init()
	' request  : characterName,oidOrder
	' response : characterName
	Dim i,j,k
	Dim sph,RS,blsResult : blsResult = False
	Dim RS2

	' preparing request parameters
	Dim characterName,oidOrder
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	characterName	= Parameters(0)
	oidOrder		= Parameters(1)

	Dim accountName
	accountName = GetAccountName(characterName)
	'accountName = "indralig"

	If accountName="" OR IsEmpty(accountName) OR IsNull(accountName) Then Call Error(99)

	' Request To ShopDB for OrderInfo
	Dim strUserID,TID,nOrderStatus,nTotalPrice,dtOrder
	Dim nReceiveQuantity

	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.public_Order_GetInfo"
		Call .InitCommand()
		Call .AppendParam("@oidOrder",adInteger,adParamInput,,oidOrder)
		blsResult = .ExecRecordset2()
	End with

	If blsResult Then
		Set RS = sph.rs
		Set RS2 = sph.rs2
		If RS.Eof Then Call Error(1)
		If RS2.Eof Then Call Error(1)
	Else
		Call Error(1)
	End If
	set sph = Nothing
	
	TID			= RS("TID")	
	strUserID	= RS("strUserID")
	nOrderStatus= RS("nOrderStatus")
	nTotalPrice	= RS("nCash")
	dtOrder		= RS("dtOrder")
	
	RS.Close

	'Response.Write strUserID &"<br />"
	'Response.Write TID &"<br />"

	If nOrderStatus<>0 Then RS2.Close:Call Error(2)
	If nTotalPrice<=0 Then RS2.Close:Call Error(101)

	Do Until RS2.Eof
		
		nReceiveQuantity = RS2("nReceiveQuantity")
		If nReceiveQuantity>0 Then RS2.Close:Call Error(102)

		RS2.MoveNext
	Loop
	RS2.Close

	If DateDiff("d",dtOrder,Now())>7 Then Call Error(103)

	' Billing
	Dim oPublisher,Result
	Set oPublisher = new Publisher
	oPublisher.DEBUG = False
	If oPublisher.refund(strUserID,TID,nTotalPrice) Then
		' Request To ShopDB for OrderCancel
		blsResult = False
		Set sph = new SPHelper
		with sph
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = cashDBconnectionString
			.SPName = "dbo.public_Order_Cancel"
			Call .InitCommand()
			Call .AppendParam("@oidOrder",adInteger,adParamInput,,oidOrder)
			Call .AppendParam("@strReason",adVarWChar,adParamInput,100,"User Cancel")
			blsResult = .ExecNoRecords()
		End with
		set sph = Nothing

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
	Else
		Call Error(3)
	End If

	Call Ok(characterName & SEPARATOR & accountName & SEPARATOR & oidOrder)
%>