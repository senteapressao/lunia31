<!--#include file="../common.asp"-->
<%
	' request  : character name, quest hash
	' response : quest hash
	
	Call Init()
	Dim characterName, questHash
	charactername=Parameters(0)
	questHash=Parameters(1)
	
	Dim params, ret
	params=Array( charactername, questHash )
	ret=ExecSP(characterDBconnectionString, "DropQuest", params)

	' return string
	Call Ok(questHash)
%>