<!--#include file="../common.asp"-->
<%
	' request  : character name, quest hash, initial state, expiredDate, select item index
	' response : quest hash, select item index, expiredDate
	
	Call Init()
	Dim characterName, questHash, initState, selectItemIndex, expiredDate
	charactername=Parameters(0)
	questHash=Parameters(1)
	initState=Parameters(2)
	expiredDate=Parameters(3)
	If UBound(Parameters)>3 Then selectItemIndex=Parameters(4)
	
	Dim params, ret
	params=Array( charactername, questHash, initState, expiredDate )
	ret=ExecSP(characterDBconnectionString, "AcceptQuest", params)
	
	If ret=0 Then
		Call Ok(questHash & SEPARATOR & selectItemIndex & SEPARATOR & expiredDate)
	Else
		Call Error(ret)
	End If
%>
