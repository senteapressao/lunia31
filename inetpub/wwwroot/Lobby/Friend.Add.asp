<!--#include file="../common.asp"-->
<%
	' request  : characterName, targetCharacterName, isFriend
	' response : 

	Call Init()

	Dim characterName, targetCharacterName, isFriend
	characterName		= Parameters(0)
	targetCharacterName	= Parameters(1)
	isFriend			= Parameters(2)

	' call Stored procedure	
	Dim params, ret
	
	params=Array(characterName, targetCharacterName, isFriend)
	
	ret=ExecSP(characterDBconnectionString, "FriendAdd", params)

	If ret=0 Then
		Call Ok(targetCharacterName & SEPARATOR & isFriend)
	Else
		Call Error(ret)
	End If
	
%>
