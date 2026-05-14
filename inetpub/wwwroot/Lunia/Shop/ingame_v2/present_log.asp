<%	Option Explicit %>
<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/common.asp"-->
<%
	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim retString : retString = ""	
	Dim strUserID
	Dim isSend,nPageSize,nPageNo
	Dim rCodePage
	Dim totaloidProduct
	Dim nRowCount,nTotalRowCount
	
	If Ubound(Parameters)<4 Then Call Error("not enough parameter")
	
	strUserID	= Parameters(0)
	isSend		= Parameters(1)
	nPageSize	= Parameters(2)
	nPageNo		= Parameters(3)
	rCodePage	= Parameters(4)

	If strUserID="" Then Call Error("no strUserID")
	If isSend="" Then Call Error("no isSend")
	
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
		.SPName = "dbo.public_Order_Present_GetList"
		Call .InitCommand()
		Call .AppendParam("@oidUser",adInteger,adParamInput,,0)
		Call .AppendParam("@strUserID",adVarChar,adParamInput,50,strUserID)
		Call .AppendParam("@isSend",adUnsignedTinyInt,adParamInput,,isSend)
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
	
	retString = retString & isSend & SEPARATOR & nTotalRowCount & SEPARATOR & nPageNo

	Dim no
	no = nTotalRowCount - (nPageSize * (nPageNo-1) )

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
		retString = retString & SEPARATOR & no &_
			SUBSEPARATOR & FormatDt(RS("dtOrder"),"SQL_TM") &_
			SUBSEPARATOR & RS("strUserID") &_
			SUBSEPARATOR & RS("strUserID_R")
		If rCodePage <> "" Then
			retString = retString & SUBSEPARATOR & arrStrProduct(rCodePage &"_"& RS("oidProduct") &"_strProduct")
		Else
			retString = retString & SUBSEPARATOR & RS("strProduct")
		End If
		retString = retString & SUBSEPARATOR & RS("strMessage")

		no = no - 1 
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
			Case "SQL_TM"
				FormatDt = Year(argDt) & "-" & Right("0" & Month(argDt), 2) & "-" & Right("0" & Day(argDt),2) & " "
				FormatDt = FormatDt & Right("0" & Hour(argDt), 2) & ":" & Right("0" & Minute(argDt), 2) & ":" & Right("0" & Second(argDt), 2)
			Case Else
				FormatDt = ""
		End Select
	End Function
%>