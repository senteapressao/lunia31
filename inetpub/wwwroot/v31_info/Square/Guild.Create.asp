<!--#include file="../common.asp"-->
<%
	' request  : guild id, guild name, character name, intro message
	' response : guild oid
	
	Call Init()
	
	Dim soap, result
	Dim guildId, guildName, characterName, introMessage, accountName
	Dim oid
	
	guildId		= Parameters(0)
	guildName	= Parameters(1)
	characterName=Parameters(2)
	introMessage= Parameters(3)
	
	accountName = GetAccountName( characterName )
	
	Set soap=Server.CreateObject("MSSOAP.SoapClient30")
	soap.ClientProperty("ServerHTTPRequest") = True
	Call soap.MSSoapInit(GuildWSDL)
	
	'Response.Write ( ServerGroupCode & " " & guildId & " " & guildName & " " & accountName & " " & characterName & " " & introMessage & " ")
	'Response.End
	result=soap.CreateGuild( ServerGroupCode, guildId, guildName, accountName, characterName, introMessage, oid )
	
	If result=0 Then ' success
		Call Ok(oid)
	Else
		If result=100 Then Call Error(1) ' invalid master id
		If result=200 Then Call Error(2) ' failed to create guild
		Call Error(99) ' unknown
	End If
%>
