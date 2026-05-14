<!--#include file="../common.asp"-->
<%
	' request  : guildId, guildMemberId, targetName, targetGradeTo
	' response : 

	Call Init()
	If UBound(parameters)<3 Then Call Error("not enough parameter")
	
	Dim guildId, guildMemberId, targetName, targetGradeTo

	guildId			= Parameters(0)
	guildMemberId	= Parameters(1)
	targetName		= Parameters(2)
	targetGradeTo	= Parameters(3)

	Dim params, ret
	params=Array( guildId, guildMemberId, targetName, targetGradeTo )
	ret=ExecSP(guildDBconnectionString, "GuildChangeGrade", params)

	If ret=0 Then
		Call Ok(Array(targetName, targetGradeTo))
	Else
		Call Error(ret)
	End If
%>
