<%	Option Explicit %>
<!--#include virtual="include/connect.asp"-->
<!--#include virtual="include/common.asp"-->
<%
	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim retString : retString = ""	
	Dim oidProduct
	Dim rCodePage
	If Ubound(Parameters)<1 Then Call Error("not enough parameter")
	
	oidProduct	= Parameters(0)
	rCodePage	= Parameters(1)

	If oidProduct="" Then Call Error("no oidProduct")
	
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
		Dim strProductName,strDescription
		strProductName	= RS("strProductName")
		strDescription	= RS("strDescription")
		'If rCodePage<>"" Then
		'	Set sph = new SPHelper
		'	with sph
		'		.DEBUG = False
		'		Set .cmd = Command
		'		.ConnStr = ConnStrShop
		'		.SPName = "dbo.public_Product_GetInfo_locale"
		'		Call .InitCommand()
		'		Call .AppendParam("@oidProduct",adInteger,adParamInput,,oidProduct)
		'		Call .AppendParam("@codePage",adVarChar,adParamInput,10,rCodePage)
		'		Call .AppendParam("@strProductName",adVarChar,adParamInputOutput,100,null)
		'		Call .AppendParam("@strDescription",adVarChar,adParamInputOutput,1000,null)
		'		blsResult = .ExecNoRecords()
		'	End with
		'
		'	If blsResult Then
		'		strProductName		= sph.GetParamValue("@strProductName")
		'		strDescription		= sph.GetParamValue("@strDescription")
		'	Else
		'		' Println sph.frk_n4ErrorCode &" / "& sph.frk_strErrorText
		'	End If
		'	set sph = Nothing
		'End If
			
		retString = retString & RS("oidProduct") &_
			SEPARATOR & RS("nCategory1") &_
			SEPARATOR & RS("nCategory2") &_
			SEPARATOR & RS("nProductStatus") &_
			SEPARATOR & RS("strProductCode") &_
			SEPARATOR & strProductName &_
			SEPARATOR & RS("nProductExpire") &_
			SEPARATOR & RS("nProductEA") &_
			SEPARATOR & strDescription &_
			SEPARATOR & RS("nProductPrice") &_
			SEPARATOR & RS("nSalePrice") &_
			SEPARATOR & RS("isUsePresent")
		
		If Not(RS2.Eof) Then
			Do Until RS2.Eof
				retString = retString & SEPARATOR & RS2("oidProduct") &_
					SUBSEPARATOR & RS2("nProductExpire") &_
					SUBSEPARATOR & RS2("nProductEA") &_
					SUBSEPARATOR & RS2("nProductPrice") &_
					SUBSEPARATOR & RS2("nSalePrice")

				RS2.MoveNext : i=i+1
			Loop
			RS2.Close
		End If
	End If
	RS.Close

	Call Ok(retString)
%>