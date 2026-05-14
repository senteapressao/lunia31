<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	If UBound(parameters)<3 Then Call Error("not enough parameter")

	' request  : character id, order number, product number, quantity
	' response : character id, order number, product number, quantity
	
	Dim characterName, orderNumber, productNumber, quantity
	characterName=Parameters(0)
	orderNumber=Parameters(1)
	productNumber=Parameters(2)
	quantity=Parameters(3)

	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim oidOrder,oidProduct,nReceiveQuantity,strReason
	
	oidOrder	= orderNumber
	oidProduct	= productNumber
	nReceiveQuantity	= quantity
	strReason	= "(game server) ["& ServerGroupName &"]"& characterName &" received."
    
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.public_Order_ProductReceive"
		Call .InitCommand()
		Call .AppendParam("@oidOrder",adInteger,adParamInput,,oidOrder)
		Call .AppendParam("@oidProduct",adInteger,adParamInput,,oidProduct)
		Call .AppendParam("@nReceiveQuantity",adSmallInt,adParamInput,,nReceiveQuantity)
		Call .AppendParam("@strReason",adVarChar,adParamInput,100,strReason)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Response.Write characterName
		Response.Write SEPARATOR & orderNumber
		Response.Write SEPARATOR & productNumber
		Response.Write SEPARATOR & quantity
	Else
		Error(sph.frk_strErrorText)
	End If
	set sph = Nothing
%>
