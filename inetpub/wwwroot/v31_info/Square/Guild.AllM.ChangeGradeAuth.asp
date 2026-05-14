<!--#include file="../common.asp"-->
<%
	' request  : guildId, grade, guildMemberId, gradeAuth
	' response : 

	Call Init()
	If UBound(parameters)<3 Then Call Error("not enough parameter")
	
	Dim guildId, grade, guildMemberId, gradeAuth

	guildId			= Parameters(0)
	grade			= Parameters(1)
	guildMemberId	= Parameters(2)
	gradeAuth		= Parameters(3)

	Dim params, ret
	params=Array( guildId, grade, guildMemberId, gradeAuth )
	ret=ExecSP(guildDBconnectionString, "GuildChangeGradeAuth", params)

	If ret=0 Then
		Call Ok(Array(grade, gradeAuth))
	Else
		Call Error(ret)
	End If
%>
