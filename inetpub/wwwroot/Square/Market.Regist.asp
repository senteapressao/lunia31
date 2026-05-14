<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  :(GET) characterName
	'			(POST)
	'			characterName,category,itemType
	'			,itemHash,stackedCount,itemSerial,instance
	'			,itemName,lighting,reinforce
	'			,petName,petLevel,isRare,rareProbability
	'			,salePrice,expireDate,regDate,showLevel
	'			,maxItemCount
	' response : characterName,marketId

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	Dim rowData
	rowData=getConvAsciiIgnore(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)

	' preparing request parameters
	Dim characterName,category,itemType _
		,itemHash,stackedCount,itemSerial,instance,itemExpire _
		,itemName,lighting,reinforce _
		,petName,petLevel,isRare,rareProbability,enchantSerial _
		,salePrice,expireDate,regDate,showLevel _
		,maxItemCount,feeDiscount
	If UBound(parameters)<21 Then Call Error("not enough parameter")
	characterName	= Parameters(0)
	category		= Parameters(1)
	itemType		= Parameters(2)
	itemHash		= Parameters(3)
	stackedCount	= Parameters(4)
	itemSerial		= Parameters(5)
	instance		= Parameters(6)
	itemExpire		= Parameters(7)
	itemName		= Parameters(8)
	lighting		= Parameters(9)
	reinforce		= Parameters(10)
	petName			= Parameters(11)
	petLevel		= Parameters(12)
	isRare			= Parameters(13)
	rareProbability	= Parameters(14)
	enchantSerial	= Parameters(15)
	salePrice		= Parameters(16)
	expireDate		= Parameters(17)
	regDate			= Parameters(18)
	showLevel		= Parameters(19)
	maxItemCount	= Parameters(20)
	feeDiscount	= Parameters(21)

	' for i = 0 to UBound(Parameters) 
	' 	Response.write("Parameters[" & i & "] " & Parameters(i) & "<br>")
		
	' Next
	
	Dim marketId

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.Market_Regist"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@category",adVarChar,adParamInput,20,category)
		Call .AppendParam("@itemType",adUnsignedTinyInt,adParamInput,,itemType)
		Call .AppendParam("@itemHash",adInteger,adParamInput,,itemHash)
		Call .AppendParam("@stackedCount",adInteger,adParamInput,,stackedCount)
		Call .AppendParam("@itemSerial",adBigInt,adParamInput,,itemSerial)
		Call .AppendParam("@instance",adBigInt,adParamInput,,instance)
		Call .AppendParam("@itemExpire",adDBTimeStamp,adParamInput,,FormatDt(itemExpire,"SQL_TM"))
		Call .AppendParam("@itemName",adVarWChar,adParamInput,250,itemName)
		Call .AppendParam("@lighting",adUnsignedTinyInt,adParamInput,,lighting)
		Call .AppendParam("@reinforce",adUnsignedTinyInt,adParamInput,,reinforce)		
		Call .AppendParam("@petName",adVarWChar,adParamInput,50,petName)		
		Call .AppendParam("@petLevel",adSmallInt,adParamInput,,petLevel)
		Call .AppendParam("@isRare",adBoolean,adParamInput,,isRare)
		Call .AppendParam("@rareProbability",adInteger,adParamInput,,rareProbability)	
		Call .AppendParam("@enchantSerial",adSingle,adParamInput,,enchantSerial)		
		Call .AppendParam("@salePrice",adBigInt,adParamInput,,salePrice)
		Call .AppendParam("@expireDate",adDBTimeStamp,adParamInput,,FormatDt(expireDate,"SQL_TM"))
		If IsDate(regDate) Then
			Call .AppendParam("@regDate",adDBTimeStamp,adParamInput,,FormatDt(regDate,"SQL_TM"))
		Else
			Call .AppendParam("@regDate",adDBTimeStamp,adParamInput,,null)
		End If
		Call .AppendParam("@showLevel",adUnsignedTinyInt,adParamInput,,showLevel)
		Call .AppendParam("@maxItemCount",adInteger,adParamInput,,maxItemCount)
		Call .AppendParam("@feeDiscount",adDouble,adParamInput,,feeDiscount)
		Call .AppendParam("@marketId",adBigInt,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		marketId = sphn.GetParamValue("@marketId")
	End If
	set sphn = Nothing

	retString = characterName & SEPARATOR & marketId

	Call Ok(retString)
%>