<!--#include file="../common.asp"-->
<%
	Call Init()
	' request  : address, port, capacity
	' response : nothing
	
	' preparing request parameters
	Dim characterName, characterStageFlags

	characterName=Parameters(0)

	' call Stored procedure	
	Dim params, ret
	params=Array(characterName)
	
	ret=ExecSP(characterDBconnectionString, "GetCharacterStateFlags", params)
	
	characterStageFlags=params(0)
	
	If ret =0 Then
		Call Ok(params(0))
	Else
		Call Error(-1)
	End If
%>
