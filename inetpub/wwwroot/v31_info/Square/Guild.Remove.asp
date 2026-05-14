<!--#include file="../common.asp"-->
<%
	' request  : guild oid
	' response : N/A
	
	Call Init()
	
	Dim soap, result
	Dim guildOid
	
	guildOid	= Parameters(0)
	
	Set soap=Server.CreateObject("MSSOAP.SoapClient30")
	soap.ClientProperty("ServerHTTPRequest") = True
	Call soap.MSSoapInit(GuildWSDL)
	
	result=soap.RemoveGuild( ServerGroupCode, guildOid, True )
	
	If result=0 Then ' success
		Call Ok(0)
	Else
		Call Error( result - 298 ) ' failed to remove guild 
	End If
%>