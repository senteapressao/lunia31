<!--#include file="../common.asp"-->
<%
	' request  : guild id, guild name, character name, intro message, stageLevel, pvpLevel, exp
	' response : guild oid
	
	Call Init()
	
	Dim soap, result
	Dim guildId, guildName, characterName, introMessage, accountName, stageLevel, pvpLevel, exp
	Dim oid
	
	guildId		= Parameters(0)
	guildName	= Parameters(1)
	characterName=Parameters(2)
	introMessage= Parameters(3)
	stageLevel	= Parameters(4)
	pvpLevel	= Parameters(5)
	exp			= Parameters(6)
	
	accountName = GetAccountName( characterName )
	
	Set soap=Server.CreateObject("MSSOAP.SoapClient30")
	soap.ClientProperty("ServerHTTPRequest") = True
	Call soap.MSSoapInit(GuildWSDL)
	
	result = soap.CreateGuildEx( ServerGroupCode, guildId, guildName, accountName, characterName, introMessage, stageLevel, pvpLevel, exp, oid )
	
	If result=0 Then ' success
		Call Ok(oid)
	Else
		If result=100 Then Call Error(1) ' invalid master id
		Call Error(result - 198)
	End If
%>
