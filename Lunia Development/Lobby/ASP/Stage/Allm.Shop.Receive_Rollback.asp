<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	If UBound(parameters)<3 Then Call Error("not enough parameter")

	' request  : character id, order number, product number. reason
	' response : none
	
	Dim characterName, orderNumber, productNumber, reason
	characterName=Parameters(0)
	orderNumber=Parameters(1)
	productNumber=Parameters(2)
	reason=Parameters(3)

	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim oidOrder,oidProduct,strReason
	
	oidOrder	= orderNumber
	oidProduct	= productNumber
	strReason	= "(game server) rollback -> "& reason
    
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.public_Order_ProductReceive_Rollback"
		Call .InitCommand()
		Call .AppendParam("@oidOrder",adInteger,adParamInput,,oidOrder)
		Call .AppendParam("@oidProduct",adInteger,adParamInput,,oidProduct)
		Call .AppendParam("@strReason",adVarChar,adParamInput,100,strReason)
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Response.Write "ok"
	Else
		Error(sph.frk_strErrorText)
	End If
	set sph = Nothing
%>