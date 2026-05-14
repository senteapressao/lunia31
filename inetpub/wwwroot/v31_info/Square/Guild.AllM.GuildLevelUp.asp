<!--#include file="../common.asp"-->
<%
	' request  : guildId, guildMemberId, guildLevel
	' response : 

	Call Init()
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	
	Dim guildId, guildMemberId, guildLevel

	guildId			= Parameters(0)
	guildMemberId	= Parameters(1)
	guildLevel		= Parameters(2)

	Dim params, ret
	params=Array( guildId, guildMemberId, guildLevel )
	ret=ExecSP(guildDBconnectionString, "GuildLevelUp", params)

	If ret=0 Then
		Call Ok(guildLevel)
	Else
		Call Error(ret)
	End If
%>
