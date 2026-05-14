<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init() 
	Dim hash,characterName,params,achieverList
	hash			= Parameters(0)
	characterName	= Parameters(1)
	params = ""'"3050844#3974365#4897886#5821407#6744928#7668449#8591970#9515491#10439012#30756474#8556010#9479531#10403052#11326573#12250094#13173615#14097136#15020657#15944178#36261640#45715458#45715459#46638979#46638980#47562500#48486021#50333063#51256584#7574046#7574047"
	achieverList = "I fucking missed that shit adsdas [br][br]|100"
	Call Ok(hash & SEPARATOR & characterName & SEPARATOR & params & SEPARATOR & "0" & SEPARATOR & achieverList)
%>
