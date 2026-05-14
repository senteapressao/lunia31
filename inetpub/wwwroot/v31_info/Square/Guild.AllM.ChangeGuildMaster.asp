<!--#include file="../common.asp"-->
<%
	' request  : guildId, guildMemberId, targetName
	' response : 

	Call Init()
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	
	Dim guildId, guildMemberId, targetName

	guildId			= Parameters(0)
	guildMemberId	= Parameters(1)
	targetName		= Parameters(2)

	Dim params, ret
	params=Array( guildId, guildMemberId, targetName )
	ret=ExecSP(guildDBconnectionString, "GuildChangeGuildMaster", params)

	If ret=0 Then
		Call Ok(targetName)
	Else
		Call Error(ret)
	End If
%>
