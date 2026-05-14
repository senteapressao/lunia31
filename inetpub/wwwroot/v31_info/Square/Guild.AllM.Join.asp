<!--#include file="../common.asp"-->
<%
	' request  : guildId, characterName, grade
	' response : 
	
	Call Init()
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	
	Dim guildId, characterName, grade

	guildId			= Parameters(0)
	characterName	= Parameters(1)
	grade			= Parameters(2)

	' return error if trying to join a user as guildmaster
	If grade=0 Then Call Error(4)

	Dim params, ret
	params=Array( guildId, characterName, grade)
	ret=ExecSP(guildDBconnectionString, "GuildJoin", params)

	If ret=0 Then
		Call Ok(guildId)
	Else
		Call Error(ret)
	End If
%>
