<!--#include file="../common.asp"-->
<%
	Call Init()
	' request  : address, port, capacity
	' response : nothing
	
	' preparing request parameters
	Dim accountName, characterName, instantStageFlags

	accountName=Parameters(0)
	characterName=Parameters(1)
	instantStageFlags=Parameters(2)
	

	' call Stored procedure	
	Dim params, ret
	params=Array(characterName, instantStageFlags)
	
	ret=ExecSP(characterDBconnectionString, "UpdateInstantStateFlags", params)
	If ret=0 Then
		Call Ok("Ok")
	Else
		Call Error(ret)
	End If
%>
