<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	If UBound(parameters)<3 Then Call Error("not enough parameter")

	' request  : character id, page number, row per page, is present, language Code
	' response : character id, account id, total row count
	'				[, packages ( separated : order number, product number, product name, quantity, representative item id { items ( subseparated : item id, expire day, stack count, description, primary option, secondary option) } ) ]
	
	Dim characterName, pageNumber, rowPerPage, isPresent, codePage
	characterName=Parameters(0)
	pageNumber=Parameters(1)
	rowPerPage=Parameters(2)
	isPresent=Parameters(3)
	codePage=Parameters(4)
	

	
	If codePage="MY_Chinese" Then
		codePage="zh-cn"
	Else
		codePage=""
	End If

	' check validation
	If (pageNumber<1) Then Call Error("invalid page index, must be larger than 0")

	Dim accountName

	accountName=GetAccountName(characterName)

	Dim RS,RS2

	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim strUserID,nPageSize,nPageNo
	Dim nRowCount,nTotalRowCount
	Dim retString : retString=""
	
	strUserID	= accountName
	nPageSize	= rowPerPage
	nPageNo		= pageNumber

	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.public_Order_ProductReceive_GetList"
		Call .InitCommand()
		Call .AppendParam("@strUserID",adVarChar,adParamInput,50,strUserID)
		Call .AppendParam("@isPresent",adUnsignedTinyInt,adParamInput,,IsPresent)
		'Call .AppendParam("@codePage",adVarChar,adParamInput,20,codePage)
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
	End If
	set sph = Nothing
	If nTotalRowCount="" OR IsNull(nTotalRowCount) Then nTotalRowCount=0

	retString = retString & characterName
	retString = retString & SEPARATOR & accountName
	retString = retString & SEPARATOR & nTotalRowCount

	If blsResult=False Then Response.End

	Do Until RS.Eof
		retString = retString & SEPARATOR & RS("oidOrder")
		retString = retString & SEPARATOR & RS("oidProduct")
		retString = retString & SEPARATOR & RS("strProductName")
		retString = retString & SEPARATOR & RS("nRemainQuantity")
		retString = retString & SEPARATOR & RS("strProductCode")
		If RS("isPackage")<>1 Then
			retString = retString & SEPARATOR & RS("strProductCode")
			retString = retString & SUBSEPARATOR & RS("nProductExpire")
			retString = retString & SUBSEPARATOR & RS("nProductEA")
			retString = retString & SUBSEPARATOR & ""
			retString = retString & SUBSEPARATOR & RS("strOption1Code")
			retString = retString & SUBSEPARATOR & RS("strOption2Code")
		Else
			RS2.Filter = "oidOrder="& RS("oidOrder") &" AND oidProduct_Package="& RS("oidProduct")
			If RS2.Eof Then Call Error("invalid package item infomation")
			i=0
			Do Until RS2.Eof
				If i=0 Then
					retString = retString & SEPARATOR
				Else
					retString = retString & SUBSEPARATOR
				End If
				retString = retString & RS2("strProductCode")
				retString = retString & SUBSEPARATOR & RS2("nProductExpire")
				retString = retString & SUBSEPARATOR & RS2("nProductEA")
				retString = retString & SUBSEPARATOR & ""
				retString = retString & SUBSEPARATOR & ""
				retString = retString & SUBSEPARATOR & ""

				RS2.MoveNext : i=i+1
			Loop
			RS2.Filter = adFilternone
		End If

		RS.MoveNext
	Loop
	RS2.Close
	RS.Close

	Response.Write retString
%>