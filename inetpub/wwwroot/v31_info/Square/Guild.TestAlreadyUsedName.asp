<!--#include file="../common.asp"-->
<%
	' request  : guild name
	' response : 0 if the guild name is already used or

	Call Init()
	
	Dim guildName, soap, result
	
	guildName=Parameters(0)
	
	Set soap=Server.CreateObject("MSSOAP.SoapClient30")
	soap.ClientProperty("ServerHTTPRequest") = True
	Call soap.MSSoapInit(GuildWSDL)
	
	result=soap.IsAlreadyUsedName(guildName)
	Set soap=Nothing
	
	If result Then
		Call Ok(1) ' it's already used guild name
	Else
		Call Ok(0)
	End If	

	' no error because it could make bulky error logs
%>