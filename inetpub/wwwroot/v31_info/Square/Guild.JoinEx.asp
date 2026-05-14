<!--#include file="../common.asp"-->
<%
	' request  : guild oid, character name, stageLevel, pvpLevel, exp
	' response : guild oid
	
	Call Init()
	
	Dim soap, result
	Dim guildOid, characterName, accountName, stageLevel, pvpLevel, exp
	
	guildOid	= Parameters(0)
	characterName=Parameters(1)
	accountName	= GetAccountName(characterName)
	stageLevel	= Parameters(2)
	pvpLevel	= Parameters(3)
	exp			= Parameters(4)
	
	Set soap=Server.CreateObject("MSSOAP.SoapClient30")
	soap.ClientProperty("ServerHTTPRequest") = True
	Call soap.MSSoapInit(GuildWSDL)
	
	result=soap.JoinGuildEx( ServerGroupCode, guildOid, accountName, characterName, stageLevel, pvpLevel, exp )
	
	If result=0 Then ' success
		Call Ok(guildOid)
	Else
		Call Error(result - 398) ' failed to join(result:400~408), error 1 occrus in game logic
	End If
%>
