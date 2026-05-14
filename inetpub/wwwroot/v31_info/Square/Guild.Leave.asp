<!--#include file="../common.asp"-->
<%
	' request  : guild oid, character name, kicked
	' response : kicked
	
	Call Init()
	
	Dim soap, result
	Dim guildOid, characterName, kicked, accountName
	
	guildOid	= Parameters(0)
	characterName=Parameters(1)
	kicked		= Parameters(2)
	accountName	= GetAccountName(characterName)
	
	Set soap=Server.CreateObject("MSSOAP.SoapClient30")
	soap.ClientProperty("ServerHTTPRequest") = True
	Call soap.MSSoapInit(GuildWSDL)
	
	result=soap.LeaveGuild( ServerGroupCode, guildOid, accountName, characterName )
	
	If result Then ' success
		Call Ok(kicked)
	Else
		Call Error(1) ' failed to leave
	End If
%>
