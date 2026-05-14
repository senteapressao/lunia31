<!--#include file="../common.asp"-->
<%
	' request  : character name, quest hash, param1, param2, param3
	' response : quest hash
	
	Call Init()
	Dim characterName, questHash, param1, param2, param3
	charactername=Parameters(0)
	questHash=Parameters(1)
	param1=Parameters(2)
	param2=Parameters(3)
	param3=Parameters(4)

	If (param1=0 And param2=0 And param3=0) Then
		Call Ok(questHash)
	End If
	
	Dim params, ret
	params=Array( charactername, questHash, param1, param2, param3 )
	ret=ExecSP(characterDBconnectionString, "UpdateQuestParameter", params)
	
	If ret=0 Then
		Call Ok(questHash)
	Else
		Call Error(ret)
	End If
%>
