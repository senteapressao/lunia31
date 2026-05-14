<%
	Dim SQL,RS,Cmd,RS2
	
	Dim ConnStrShop
	ConnStrShop = GetConnStr("DESKTOP-S29QIQ5\SQLEXPRESS","d-shop","sa","algoraSaturn2203@")
	Dim ConnStrCharacter
	ConnStrCharacter = GetConnStr("DESKTOP-S29QIQ5\SQLEXPRESS","v3_character","sa","algoraSaturn2203@")
	
%>
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Connection" id="Conn" viewastext="viewastext"></object>
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<!--#include virtual="./include/db_helper.asp"-->
<!--#include virtual="./include/DBHelper_NoTran.asp"-->