<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : itemType,itemHash,lighting,reinforce,petLevel,isRare
	' response : itemType,itemHash,lighting,reinforce,petLevel,isRare,avgPrice,count

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim itemType,itemHash,lighting,reinforce,petLevel,isRare
	If UBound(parameters)<5 Then Call Error("not enough parameter")
	itemType = Parameters(0)
	itemHash = Parameters(1)
	lighting = Parameters(2)
	reinforce = Parameters(3)
	petLevel = Parameters(4)
	isRare = Parameters(5)

	Dim avgPrice,count

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.Market_GetAvgPrice"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@itemType",adUnsignedTinyInt,adParamInput,,itemType)
		Call .AppendParam("@itemHash",adInteger,adParamInput,,itemHash)
		Call .AppendParam("@lighting",adBigInt,adParamInput,,lighting)
		Call .AppendParam("@reinforce",adBigInt,adParamInput,,reinforce)
		Call .AppendParam("@petLevel",adSmallInt,adParamInput,,petLevel)
		Call .AppendParam("@isRare",adBoolean,adParamInput,,isRare)		
		Call .AppendParam("@avgPrice",adBigInt,adParamOutput,,null)
		Call .AppendParam("@count",adInteger,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		avgPrice = sphn.GetParamValue("@avgPrice")
		count = sphn.GetParamValue("@count")
	End If
	set sphn = Nothing

	retString = itemType & SEPARATOR & itemHash &_
			SEPARATOR & lighting & SEPARATOR & reinforce &_
			SEPARATOR & petLevel & SEPARATOR & isRare &_
			SEPARATOR & avgPrice & SEPARATOR & count

	Call OK(retString)
%>