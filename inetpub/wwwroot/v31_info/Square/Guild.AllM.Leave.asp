<!--#include file="../common.asp"-->
<%
	' request  : guildId, guildMemberId
	' response : 

	Call Init()
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	
	Dim guildId, guildMemberId

	guildId			= Parameters(0)
	guildMemberId	= Parameters(1)

	Dim params, ret
	params=Array( guildId, guildMemberId)
	ret=ExecSP(guildDBconnectionString, "GuildLeave", params)

	If ret=0 Then
		Call Ok("test ok")
	Else
		Call Error(ret)
	End If
%>
