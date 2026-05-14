<!--#include file="../common.asp"-->
<%
	' request  : guildId, characterName
	' response : 
	
	Call Init()
	If UBound(parameters)<1 Then Call Error("not enough parameter")

	Dim guildId, characterName

	guildId			= Parameters(0)
	characterName	= Parameters(1)

	Dim params, ret
	params=Array( guildId, characterName)
	ret=ExecSP(guildDBconnectionString, "GuildClose", params)

	If ret=0 Then
		Call Ok("test ok")
	Else
		Call Error(result)
	End If
%>
