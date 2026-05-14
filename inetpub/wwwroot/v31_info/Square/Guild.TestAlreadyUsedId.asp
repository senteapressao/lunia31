<!--#include file="../common.asp"-->
<%
	' request  : guild id
	' response : 0 if the guild id is already used or

	Call Init()
	
	Dim guildId, soap, result
	
	guildId=Parameters(0)
	
	Set soap=Server.CreateObject("MSSOAP.SoapClient30")
	soap.ClientProperty("ServerHTTPRequest") = True
	Call soap.MSSoapInit(GuildWSDL)
	
	result=soap.IsAlreadyUsedGuildId(guildId)
	Set soap=Nothing
	
	If result Then
		Call Ok(1) ' it's already used guild id
	Else
		Call Ok(0)
	End If	
	
	' no error because it could make bulky error logs
%>
