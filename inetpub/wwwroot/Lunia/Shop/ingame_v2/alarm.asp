<%	Option Explicit %>
<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/common.asp"-->
<%
	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim retString : retString = ""
	Dim strUserID
	Dim nPresentCount,nCouponCount,nCouponCount_new
		
	If Ubound(Parameters)<0 Then Call Error("not enough parameter")
	
	strUserID	= Parameters(0)
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = ConnStrShop
		.SPName = "dbo.public_Order_ProductReceive_GetList"
		Call .InitCommand()
		Call .AppendParam("@strUserID",adVarChar,adParamInput,50,strUserID)
		Call .AppendParam("@isPresent",adUnsignedTinyInt,adParamInput,,1)
		Call .AppendParam("@codePage",adVarChar,adParamInput,20,null)
		Call .AppendParam("@nPageSize",adUnsignedTinyInt,adParamInput,,5)
		Call .AppendParam("@nPageNo",adInteger,adParamInput,,1)
		Call .AppendParam("@nRowCount",adInteger,adParamInputOutput,,null)
		Call .AppendParam("@nTotalRowCount",adInteger,adParamInputOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		nPresentCount = sph.GetParamValue("@nTotalRowCount")
	Else
		nPresentCount = 0
	End If
	set sph = Nothing

	nCouponCount = 0

	blsResult = False
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = ConnStrShop
		.SPName = "dbo.public_Coupon_User_GetCount_New"
		Call .InitCommand()
		Call .AppendParam("@oidUser",adInteger,adParamInput,,null)
		Call .AppendParam("@strUserID",adVarChar,adParamInput,50,strUserID)
		Call .AppendParam("@nTotalCount",adInteger,adParamInputOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		nCouponCount_new = sph.GetParamValue("@nTotalCount")
	Else
		nCouponCount_new = 0
	End If
	set sph = Nothing
	
	Call Ok(strUserID & SEPARATOR & nPresentCount & SEPARATOR & nCouponCount & SEPARATOR & nCouponCount_new)
%>