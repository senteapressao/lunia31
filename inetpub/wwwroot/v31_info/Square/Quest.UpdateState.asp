<!--#include file="../common.asp"-->
<%
	' request  : character name, quest hash, new state
	' response : quest hash
	
	Call Init()
	Dim characterName, questHash, newState
	charactername=Parameters(0)
	questHash=Parameters(1)
	newState=Parameters(2)
	
	Dim params, ret
	params=Array( charactername, questHash, newState )
	ret=ExecSP(characterDBconnectionString, "UpdateQuestState", params)
	
	If ret=0 Then
		Call Ok(questHash)
	Else
		Call Error(ret)
	End If
%>
