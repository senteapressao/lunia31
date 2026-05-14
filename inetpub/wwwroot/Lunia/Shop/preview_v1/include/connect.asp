<%
	Dim ConnStr,ConnStrGame
	Dim SQL,RS,Cmd,RS2

	ConnStr = GetConnStr("","","","")
	ConnStrGame = GetConnStr("","","","")
	
	Dim ConnStrShop
	ConnStrShop = GetConnStr("DESKTOP-S29QIQ5\SQLEXPRESS","d-shop","sa","algoraSaturn2203@")
	Dim ConnStrShopReal
	ConnStrShopReal = GetConnStr("DESKTOP-S29QIQ5\SQLEXPRESS","shop","sa","algoraSaturn2203@")	
	Dim ConnStrShopTest
	ConnStrShopTest = GetConnStr("DESKTOP-S29QIQ5\SQLEXPRESS","d-shop","sa","algoraSaturn2203@")
	
	Dim LoginWSDL,ShopWSDL
	LoginWSDL = "http://DESKTOP-S29QIQ5\SQLEXPRESS/ServiceLogin.asmx?wsdl"
	ShopWSDL = "http://DESKTOP-S29QIQ5\SQLEXPRESS/ServiceLogin.asmx?wsdl"

%>
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Connection" id="Conn" viewastext="viewastext"></object>
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<!--#include virtual="/lib/include/db_helper.asp"-->