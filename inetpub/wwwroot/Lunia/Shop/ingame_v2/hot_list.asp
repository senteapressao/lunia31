<%	Option Explicit %>
<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/common.asp"-->
<%
	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim retString : retString = ""
	Dim rCodePage
	Dim totaloidProduct
	Dim nRowCount,nTotalRowCount
	
	If Ubound(Parameters)<0 Then Call Error("not enough parameter")
	
	rCodePage	= Parameters(0)
	
	If rCodePage = "French" Then
		rCodePage = "fr"
	ElseIf rCodePage = "MY_Chinese" Then
		rCodePage = "zh-cn"
	Else
		rCodePage = ""
	End If
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = ConnStrShop
		.SPName = "dbo.public_Product_HotSale_GetList_v2"
		Call .InitCommand()
			Call .AppendParam("@nPageSize",adUnsignedTinyInt,adParamInput,,10)
			Call .AppendParam("@nPageNo",adInteger,adParamInput,,1)
			Call .AppendParam("@nRowCount",adInteger,adParamInputOutput,,null)
			Call .AppendParam("@nTotalRowCount",adInteger,adParamInputOutput,,null)
			blsResult = .ExecRecordset2()
	End with

	If blsResult Then
		nRowCount = sph.GetParamValue("@nRowCount")
		nTotalRowCount = sph.GetParamValue("@nTotalRowCount")
		Set RS = sph.rs
		'Set RS2 = sph.rs2
	Else
		Call Error("data error")
	End If
	set sph = Nothing
	
	'only EU
	If rCodePage<>"" Then
		totaloidProduct = ""
		If Not(RS.EoF) Then
			Do Until RS.EoF
				If totaloidProduct = "" Then
					totaloidProduct = RS("oidProduct")
				Else
					totaloidProduct = totaloidProduct &","& RS("oidProduct")
				End If
				RS.MoveNext
			Loop
			RS.MoveFirst
			
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
			
				Dim arrStrProduct
				Set arrStrProduct = Server.CreateObject("Scripting.Dictionary")
				Do Until RS2.EoF
					arrStrProduct(RS2(1) &"_"& RS2(0) &"_strProduct") = RS2(2)
					RS2.MoveNext
				Loop
				RS2.Close
			Else
				' Println sph.frk_n4ErrorCode &" / "& sph.frk_strErrorText
			End If
			set sph = Nothing
		End If
	End If
	
	Do Until RS.Eof
		If rs("nCategory1")<>30 Then
		If retString<>"" Then retString = retString & SEPARATOR
		retString = retString & RS("oidProduct")
		If rCodePage<>"" Then
			retString = retString & SUBSEPARATOR & arrStrProduct(rCodePage &"_"& RS("oidProduct") &"_strProduct")
		Else
			retString = retString & SUBSEPARATOR & RS("strProductName")
		End If
		End If

		RS.MoveNext
	Loop
	'RS2.Close
	RS.Close

	Call Ok(retString)
%>