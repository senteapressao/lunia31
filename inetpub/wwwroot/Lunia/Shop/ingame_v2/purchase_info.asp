<%	Option Explicit %>
<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/common.asp"-->
<%
	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim retString : retString = ""	
	Dim strUserID,oidProduct
	Dim rCodePage
	Dim totaloidProduct
	
	If Ubound(Parameters)<2 Then Call Error("not enough parameter")
	
	strUserID	= Parameters(0)
	oidProduct	= Parameters(1)
	rCodePage	= Parameters(2)

	If strUserID="" Then Call Error("no strUserID")
	If oidProduct="" Then Call Error("no oidProduct")
	
	If rCodePage = "French" Then
		rCodePage = "fr"
	ElseIf rCodePage = "MY_Chinese" Then
		rCodePage = "zh-cn"
	Else
		rCodePage = ""
	End If
	
	Dim nCategory1,nCategory2
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = ConnStrShop
		.SPName = "dbo.public_Product_GetInfo"
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
		
		retString = retString & RS("oidProduct") &_
			SEPARATOR & RS("nCategory1") &_
			SEPARATOR & RS("nCategory2") &_
			SEPARATOR & RS("strProductCode") &_
			SEPARATOR & RS("strProductName") &_
			SEPARATOR & RS("nProductExpire") &_
			SEPARATOR & RS("nProductEA") &_			
			SEPARATOR & RS("nProductPrice") &_
			SEPARATOR & RS("nSalePrice") &_
			SEPARATOR & RS("nMaxOrderQuantity") &_
			SEPARATOR & RS("isUsePresent")
		
		If Not(RS2.Eof) Then
			Do Until RS2.Eof
				retString = retString & SEPARATOR & RS2("oidProduct") &_
					SUBSEPARATOR & RS2("nProductExpire") &_
					SUBSEPARATOR & 11 &_
					SUBSEPARATOR & RS2("nProductPrice") &_
					SUBSEPARATOR & RS2("nSalePrice") &_
					SUBSEPARATOR & RS2("nMaxOrderQuantity")

				RS2.MoveNext : i=i+1
			Loop
			RS2.Close
		End If
		
		nCategory1 = RS("nCategory1")
		nCategory2 = RS("nCategory2")
	End If
	RS.Close
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = ConnStrShop
		.SPName = "dbo.public_Coupon_User_GetList_Extend"
		Call .InitCommand()
		Call .AppendParam("@oidUser",adInteger,adParamInput,,null)
		Call .AppendParam("@strUserID",adVarChar,adParamInput,50,strUserID)
		Call .AppendParam("@nCategory1",adInteger,adParamInput,,nCategory1)
		Call .AppendParam("@nCategory2",adInteger,adParamInput,,nCategory2)
		Call .AppendParam("@oidProduct",adInteger,adParamInput,,oidProduct)
		blsResult = .ExecRecordset()
	End with
	
	If blsResult Then
		Set RS = sph.rs
	Else
		Call Error("data error")
	End If
	set sph = Nothing
	
	If Not(RS.Eof) Then retString = retString & SEPARATOR & "#coupon#"
	
	If rCodePage<>"" Then
		totaloidProduct = ""
		Do Until RS.EoF
			If totaloidProduct = "" Then
				totaloidProduct = RS("oidProduct")
			Else
				totaloidProduct = totaloidProduct &","& RS("oidProduct")
			End If
			RS.MoveNext
		Loop
		
		Set sph = new SPHelper
		with sph
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = ConnStrShop
			.SPName = "dbo.public_Product_GetInfo_Extend_Locale"
			Call .InitCommand()
			Call .AppendParam("@strProduct",adVarChar,adParamInput,1000,totaloidProduct)
			Call .AppendParam("@strProduct",adVarChar,adParamInput,10,rCodePage)
			blsResult = .ExecRecordset()
		End with

		If blsResult Then
			Set RS2 = sph.rs
		
			Dim arrStrTarget
			Set arrStrTarget = Server.CreateObject("Scripting.Dictionary")
			Do Until RS2.EoF
				arrStrTarget(RS2(1) &"_"& RS2(0) &"_strTarget") = RS2(2)
				RS2.MoveNext
			Loop
			RS2.Close
			
			RS.MoveFirst
		Else
			' Println sph.frk_n4ErrorCode &" / "& sph.frk_strErrorText
		End If
		set sph = Nothing
	End If
	
	Do Until RS.Eof
		retString = retString & SEPARATOR & RS("oidCoupon_User") &_
			SUBSEPARATOR & RS("strCoupon") &_
			SUBSEPARATOR & RS("nType") &_
			SUBSEPARATOR & RS("nCategory1") &_
			SUBSEPARATOR & RS("nCategory2") &_
			SUBSEPARATOR & RS("oidProduct")
		If rCodePage<>"" Then
			retString = retString & SUBSEPARATOR & arrStrTarget(rCodePage &"_"& RS("oidProduct") &"_strTarget")
		Else
			retString = retString & SUBSEPARATOR & RS("strTarget")
		End If
		retString = retString & SUBSEPARATOR & RS("nMethod") &_
			SUBSEPARATOR & RS("nValue") &_
			SUBSEPARATOR & FormatDt(RS("dtExpire"),"SQL")

		RS.MoveNext : i=i+1
	Loop
	RS.Close	

	Call Ok(retString)
%>
<%
	Function FormatDt(argDt,argType)
		Select Case CStr("" & argType)
			Case "SQL"
				FormatDt = Year(argDt) & "-" & Right("0" & Month(argDt), 2) & "-" & Right("0" & Day(argDt), 2)
			Case Else
				FormatDt = ""
		End Select
	End Function
%>