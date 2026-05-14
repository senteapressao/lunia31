<!--#include file="../common.asp"-->
<%
	' request  : guildName, masterName, guildMessage, grade0, grade1, grade2, grade3, grade4
	' response : 
	
	Call Init()
	If UBound(parameters)<7 Then Call Error("not enough parameter")

	Dim guildName, masterName, guildMessage, grade0, grade1, grade2, grade3, grade4

	guildName	= Parameters(0)
	masterName	= Parameters(1)
	guildMessage= Parameters(2)
	grade0		= Parameters(3)
	grade1		= Parameters(4)
	grade2		= Parameters(5)
	grade3		= Parameters(6)
	grade4		= Parameters(7)

	Dim params, ret, guildId
	params=Array( guildName, masterName, guildMessage, grade0, grade1, grade2, grade3, grade4)
	ret=ExecSP(guildDBconnectionString, "GuildCreate", params)

	If ret=0 Then
		Call Ok(params)

	Else
		If result=1 Then Call Error(1) ' already existed guildName
		If result=2 Then Call Error(2) ' existed charscterName in guildMember table
		Call Error(99) ' unknown
	End If
%>
