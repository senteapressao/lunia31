<%	Option Explicit %>
<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/common.asp"-->
<%
	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim retString : retString = ""
	Dim strUserID
	Dim nCash
	Dim GotoPage
	Dim nPageSize,nPageNo
	Dim rCodePage
	Dim totaloidProduct
	Dim nRowCount,nTotalRowCount
	
	If Ubound(Parameters)<3 Then Call Error("not enough parameter")
	
	strUserID	= Parameters(0)	
	nPageSize	= Parameters(1)
	nPageNo		= Parameters(2)
	rCodePage	= Parameters(3)
	
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
		.SPName = "dbo.public_Coupon_User_GetList"
		Call .InitCommand()
		Call .AppendParam("@oidUser",adInteger,adParamInput,,null)
		Call .AppendParam("@strUserID",adVarChar,adParamInput,50,strUserID)
		Call .AppendParam("@nPageSize",adUnsignedTinyInt,adParamInput,,nPageSize)
		Call .AppendParam("@nPageNo",adInteger,adParamInput,,nPageNo)
		Call .AppendParam("@nRowCount",adInteger,adParamInputOutput,,null)
		Call .AppendParam("@nTotalRowCount",adInteger,adParamInputOutput,,null)
		blsResult = .ExecRecordset()
	End with

	If blsResult Then
		nRowCount = sph.GetParamValue("@nRowCount")
		nTotalRowCount = sph.GetParamValue("@nTotalRowCount")
		Set RS = sph.rs
	Else
		Call Error("data error")
	End If
	set sph = Nothing
	
	retString = retString & nTotalRowCount & SEPARATOR & nPageNo

	'only EU, MY
	If rCodePage<>"" Then
		If Not(RS.EoF) Then
			totaloidProduct = ""
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
			
				Dim arrStrTarget
				Set arrStrTarget = Server.CreateObject("Scripting.Dictionary")
				Do Until RS2.EoF
					arrStrTarget(RS2(1) &"_"& RS2(0) &"_strTarget") = RS2(2)
					RS2.MoveNext
				Loop
				RS2.Close			
			Else
				' Println sph.frk_n4ErrorCode &" / "& sph.frk_strErrorText
			End If
			set sph = Nothing
		End If
	End If
	
	Do Until RS.EOF
		retString = retString & SEPARATOR & RS("oidCoupon") &_
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
		RS.MoveNext
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