<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Server.ScriptTimeOut = 15

	Call Init()
	' request  : characterName,strSerial
	' response : characterName
	Dim i,j,k
	Dim sph,RS,blsResult : blsResult = False

	' preparing request parameters
	Dim characterName,strSerial
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	characterName	= Parameters(0)
	strSerial = Parameters(1)
	
	Dim accountName
	accountName = GetAccountName(characterName)
	'accountName = "indralig"

	If accountName="" Or IsEmpty(accountName) Or IsNull(accountName) Then Call Error(99)
	
	' Request To ShopDB for ProductInfo
	Dim oidUser,strUserID
	oidUser	= 0
	strUserID = accountName
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.Public_Coupon_User_Create"		
		Call .InitCommand()
		Call .AppendParam("@oidUser",adInteger,adParamInput,,oidUser)
		Call .AppendParam("@strUserID",adVarChar,adParamInput,50,strUserID)
		Call .AppendParam("@strSerial",adVarChar,adParamInput,30,strSerial)
		Call .AppendParam("@strCoupon",adVarChar,adParamInputOutput,200,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
	Else
		If sph.frk_n4ErrorCode=11 Then
			Call Error(2)
		ElseIf sph.frk_n4ErrorCode=12 Then
			Call Error(3)
		ElseIf sph.frk_n4ErrorCode=13 Then
			Call Error(4)
		Else
			Call Error(1)
		End If
	End If
	set sph = Nothing

	Call Ok(characterName & SEPARATOR & accountName)
%>