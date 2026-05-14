<!--#include file="../common.asp"-->
<%
	' request  : guildId, kickerMemberId, characterName
	' response : 
	
	Call Init()
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	
	Dim guildId, kickerMemberId, characterName
	guildId			= Parameters(0)
	kickerMemberId	= Parameters(1)
	characterName	= Parameters(2)

	Dim params, ret
	params=Array( guildId, kickerMemberId, characterName)
	ret=ExecSP(guildDBconnectionString, "GuildKick", params)

	If ret=0 Then
		Call Ok(characterName)
	Else
		Call Error(ret)
	End If
%>
