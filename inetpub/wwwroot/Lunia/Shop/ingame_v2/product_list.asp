<%	Option Explicit %>
<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/common.asp"-->
<%
	'Parameters = Split("00101Portuguese",SEPARATOR)
	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim retString : retString = ""	
	Dim nCategory1,nCategory2,oidProduct_search,strProductName_search,nPageSize,nPageNo
	Dim rCodePage
	Dim totaloidProduct
	Dim nRowCount,nTotalRowCount
	Dim nStatus
	If Ubound(Parameters)<6 Then Call Error("not enough parameter")
	
	nCategory1	= Parameters(0)
	nCategory2	= Parameters(1)	
	oidProduct_search	= Parameters(2)
	strProductName_search	= Parameters(3)
	nPageSize	= Parameters(4)
	nPageNo		= Parameters(5)
	rCodePage	= ""
	If nCategory1="0" and nCategory2="0" then
		nCategory1 = "1"
		nCategory2 = "0"
	end if
	If nCategory1="" Then nCategory1 = null
	If nCategory2="" OR nCategory2="0" Then nCategory2 = null
	If oidProduct_search="" Then oidProduct_search = null
	If strProductName_search="" Then
		strProductName_search = null
	Else
		strProductName_search = Replace(strProductName_search,"%20"," ")
	End If
	
	' If nCategory1="0" Then
		' If nCategory2="3" Then
			' nStatus = 16
		' ElseIf nCategory2="2" Then
			' nStatus = 8
		' ElseIf nCategory2="1" Then
			' nStatus = 4
		' Else
			' nStatus = null
		' End If
		
		' nCategory1	= null
		' nCategory2	= null		
	' Else
		' nStatus = null
	' End If
	
	
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = ConnStrShop
		If rCodePage<>"" Then
			.SPName = "dbo.public_Ingame_Product_GetList_Locale"
		Else
			.SPName = "dbo.public_Ingame_Product_GetList"
		End If
		Call .InitCommand()
		Call .AppendParam("@nCategory1",adSmallint,adParamInput,,nCategory1)
		Call .AppendParam("@nCategory2",adSmallint,adParamInput,,nCategory2)
		Call .AppendParam("@nStatus",adUnsignedTinyInt,adParamInput,,nStatus)
		Call .AppendParam("@oidProduct_search",adInteger,adParamInput,,oidProduct_search)
		Call .AppendParam("@strProductName_search",adVarChar,adParamInput,100,strProductName_search)
		Call .AppendParam("@nPageSize",adUnsignedTinyInt,adParamInput,,nPageSize)
		Call .AppendParam("@nPageNo",adInteger,adParamInput,,nPageNo)
		Call .AppendParam("@nRowCount",adInteger,adParamInputOutput,,null)
		Call .AppendParam("@nTotalRowCount",adInteger,adParamInputOutput,,null)
		If rCodePage<>"" Then
			Call .AppendParam("@codePage",adVarChar,adParamInput,10,rCodePage)
		End If
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
	
	retString = retString & nTotalRowCount & SEPARATOR & nPageNo
	
	Do Until RS.Eof
			retString = retString & SEPARATOR & RS("oidProduct") &_
				SUBSEPARATOR & RS("nCategory1") &_
				SUBSEPARATOR & RS("nCategory2") &_
				SUBSEPARATOR & RS("nProductStatus") &_
				SUBSEPARATOR & RS("strProductCode") &_
				SUBSEPARATOR & RS("strProductName") &_
				SUBSEPARATOR & RS("nProductPrice") &_
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