<!--#include file="../common.asp"-->
<%
	' request  : guild oid, character name
	' response : guild oid
	
	Call Init()
	
	Dim soap, result
	Dim guildOid, characterName, accountName
	
	guildOid	= Parameters(0)
	characterName=Parameters(1)
	accountName	= GetAccountName(characterName)
	
	Set soap=Server.CreateObject("MSSOAP.SoapClient30")
	soap.ClientProperty("ServerHTTPRequest") = True
	Call soap.MSSoapInit(GuildWSDL)
	
	result=soap.JoinGuild( ServerGroupCode, guildOid, accountName, characterName )
	
	If result Then ' success
		Call Ok(guildOid)
	Else
		Call Error(1) ' failed to join
	End If
%>
