<%	Option Explicit %>
<!--#include virtual="./include/connect.asp"-->
<!--#include file="./common.asp"-->
<%
	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim retString : retString = ""	
	Dim Cat1,Cat2
	Dim nCategory1,nCategory2,nStatus,strProductName,nLowLevel,nHighLevel,nOrderType,nPageSize,nPageNo
	Dim nRowCount,nTotalRowCount
	
	If Ubound(Parameters)<7 Then Call Error("not enough parameter")
	
	Cat1	= Parameters(0)
	Cat2	= Parameters(1)	
	strProductName	= Parameters(2)
	nLowLevel	= Parameters(3)
	nHighLevel	= Parameters(4)
	nOrderType	= Parameters(5)
	nPageSize	= Parameters(6)
	nPageNo		= Parameters(7)
	
'	Println Request.QueryString
	
	Call SetCategory(Cat1,Cat2)
	
	If nCategory1="" Then nCategory1 = null
	If nCategory2="" Then nCategory2 = null
	nStatus = null
	If strProductName="" Then strProductName = null
	If nLowLevel="" Then nLowLevel = null
	If nHighLevel="" Then nHighLevel = null
	If nOrderType="" Then nOrderType = 0
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = ConnStrShop
		If nOrderType=0 Then
			.SPName = "dbo.public_Product_Preview_GetList_DateDesc"
		ElseIf nOrderType=1 Then
			.SPName = "dbo.public_Product_Preview_GetList_NameAsc"
		ElseIf nOrderType=3 Then
			.SPName = "dbo.public_Product_Preview_GetList_PriceAsc"
		Else
			.SPName = "dbo.public_Product_Preview_GetList"
		End If
		Call .InitCommand()
		Call .AppendParam("@nCategory1",adSmallint,adParamInput,,nCategory1)
		Call .AppendParam("@nCategory2",adSmallint,adParamInput,,nCategory2)
		Call .AppendParam("@nStatus",adUnsignedTinyInt,adParamInput,,nStatus)		
		Call .AppendParam("@strProductName_search",adVarWChar,adParamInput,100,strProductName)
		Call .AppendParam("@nLowLevel",adSmallInt,adParamInput,,nLowLevel)
		Call .AppendParam("@nHighLevel",adSmallInt,adParamInput,,nHighLevel)
		Call .AppendParam("@nPageSize",adUnsignedTinyInt,adParamInput,,nPageSize)
		Call .AppendParam("@nPageNo",adInteger,adParamInput,,nPageNo)
		Call .AppendParam("@nRowCount",adInteger,adParamInputOutput,,null)
		Call .AppendParam("@nTotalRowCount",adInteger,adParamInputOutput,,null)
		blsResult = .ExecRecordset2()
	End with

	If blsResult Then
		nRowCount = sph.GetParamValue("@nRowCount")
		nTotalRowCount = sph.GetParamValue("@nTotalRowCount")
		Set RS = sph.rs
		Set RS2 = sph.rs2
	Else
		Call Error("data error")
	End If
	set sph = Nothing
	
	retString = retString & SEPARATOR & nTotalRowCount & SEPARATOR & nPageNo
	
	Do Until RS.Eof
		retString = retString & SEPARATOR & RS("oidProduct") &_
			SUBSEPARATOR & RS("strProductCode") &_
			SUBSEPARATOR & RS("strProductName") &_
			SUBSEPARATOR & RS("nSalePrice") &_
			SUBSEPARATOR & RS("isUsePresent")

		If RS("isPackage")=1 Then
			RS2.Filter = "oidProduct_Package="& RS("oidProduct")
			Do Until RS2.Eof
				retString = retString &_
					SUBSEPARATOR & RS2("strProductCode")

				RS2.MoveNext : i=i+1
			Loop
			RS2.Filter = adFilternone
		End If

		RS.MoveNext
	Loop
	RS2.Close
	RS.Close

	Call Ok(retString)
%>
<%
	Sub SetCategory(Cat1,Cat2)
		If Cat1=99 Then
			nCategory1 = 12 : nCategory2 = 1
		ElseIf Cat1=98 AND Cat2=1 Then
			nCategory1 = 14 : nCategory2 = 1
		ElseIf Cat1=98 AND Cat2=2 Then
			nCategory1 = 14 : nCategory2 = 3
		ElseIf Cat1=1 AND Cat2=1 Then
			nCategory1 = 10 : nCategory2 = 1
		ElseIf Cat1=2 AND Cat2=1 Then
			nCategory1 = 10 : nCategory2 = 2
		ElseIf Cat1=3 AND Cat2=1 Then
			nCategory1 = 10 : nCategory2 = 3
		ElseIf Cat1=4 AND Cat2=1 Then
			nCategory1 = 10 : nCategory2 = 4
		ElseIf Cat1=5 AND Cat2=1 Then
			nCategory1 = 10 : nCategory2 = 5
		ElseIf Cat1=6 AND Cat2=1 Then
			nCategory1 = 10 : nCategory2 = 6
		ElseIf Cat1=7 AND Cat2=1 Then
			nCategory1 = 10 : nCategory2 = 7
		ElseIf Cat1=8 AND Cat2=1 Then
			nCategory1 = 10 : nCategory2 = 9
		ElseIf Cat1=9 AND Cat2=1 Then
			nCategory1 = 10 : nCategory2 = 10
		ElseIf Cat1=10 AND Cat2=1 Then
			nCategory1 = 10 : nCategory2 = 11
		ElseIf Cat1=11 AND Cat2=1 Then
			nCategory1 = 10 : nCategory2 = 12
		ElseIf Cat1=12 AND Cat2=1 Then
			nCategory1 = 10 : nCategory2 = 13
		Else
			Call Error("category data error")
		End If
	End Sub
%>