<!--#include file="../common.asp"-->
<%
	' request  : characterName, targetCharacterName
	' response : 

	Call Init()

	Dim characterName, targetCharacterName
	characterName		= Parameters(0)
	targetCharacterName	= Parameters(1)

	' call Stored procedure	
	Dim params, ret
	
	params=Array(characterName, targetCharacterName)
	
	ret=ExecSP(characterDBconnectionString, "FriendRemove", params)

	If ret=0 Then
		Call Ok(targetCharacterName)
	Else
		Call Error(ret)
	End If
	
%>
