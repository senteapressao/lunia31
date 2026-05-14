<%	Option Explicit %>
<!--#include virtual="./include/connect.asp"-->
<!--#include file="./common.asp"-->
<%
	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim retString : retString = ""	
	Dim nRowCount
	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = ConnStrShop
		.SPName = "dbo.public_Product_HotSale_GetList"
		Call .InitCommand()
		Call .AppendParam("@nRowCount",adInteger,adParamInputOutput,,null)
		blsResult = .ExecRecordset2()
	End with

	If blsResult Then
		nRowCount = sph.GetParamValue("@nRowCount")
		Set RS = sph.rs
		Set RS2 = sph.rs2
	Else
		Call Error("data error")
	End If
	set sph = Nothing
	
	Do Until RS.Eof
		retString = retString & SEPARATOR & RS("oidProduct") &_
			SUBSEPARATOR & RS("strProductCode") &_
			SUBSEPARATOR & RS("strProductName") &_
			SUBSEPARATOR & RS("nSalePrice") &_
			SUBSEPARATOR & RS("isUsePresent")

		RS.MoveNext
	Loop
	RS2.Close
	RS.Close

	Call Ok(retString)
%>
