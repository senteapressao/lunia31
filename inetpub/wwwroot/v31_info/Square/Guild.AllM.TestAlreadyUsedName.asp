<!--#include file="../common.asp"-->
<%
	' request  : guildName
	' response : 0 if the guild name is already used or

	Call Init()
	
	Dim guildName, result

	guildName=Parameters(0)

	result = ExecSP(guildDBconnectionString, "GuildIsAlreadyUsedName", guildName)

	'I think it is better that return 0 if there is no matching guildName
	'Because it means 'U can create new guild with this name'
	'But in nexon guild system returns 1 when such condition

	If result = 1 Then
		Call Ok(0) ' it's already used guild name
	Else
		Call Ok(1)
	End If	

	' no error because it could make bulky error logs
%>