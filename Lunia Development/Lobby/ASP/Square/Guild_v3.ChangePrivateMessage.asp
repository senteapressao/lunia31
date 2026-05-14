<!--#include file="../common.asp"-->
<%
	' request  : guildId, newMessage, guildMemberId
	' response : 

	Call Init()
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	
	Dim guildId, newMessage, guildMemberId

	guildId			= Parameters(0)
	newMessage		= Parameters(1)
	guildMemberId	= Parameters(2)

	Dim params, ret
	params=Array( guildId, newMessage, guildMemberId )
	ret=ExecSP(guildDBconnectionString, "public_ChangePrivateMessage", params)

	If ret=0 Then
		Call Ok(newMessage)
	Else
		Call Error(ret)
	End If
%>
