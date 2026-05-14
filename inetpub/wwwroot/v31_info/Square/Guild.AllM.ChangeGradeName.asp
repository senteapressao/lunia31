<!--#include file="../common.asp"-->
<%
	' request  : guildId, grade, guildMemberId, gradeName
	' response : 

	Call Init()
	If UBound(parameters)<3 Then Call Error("not enough parameter")
	
	Dim guildId, grade, guildMemberId, gradeName

	guildId			= Parameters(0)
	grade			= Parameters(1)
	guildMemberId	= Parameters(2)
	gradeName		= Parameters(3)

	Dim params, ret
	params=Array( guildId, grade, guildMemberId, gradeName )
	ret=ExecSP(guildDBconnectionString, "GuildChangeGradeName", params)

	If ret=0 Then
		Call Ok(Array(grade, gradeName))
	Else
		Call Error(ret)
	End If
%>
