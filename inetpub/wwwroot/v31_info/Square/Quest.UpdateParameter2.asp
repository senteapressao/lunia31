<!--#include file="../common.asp"-->
<%
	' request  : character name, { quest hash, param1, param2, param3 }
	' response : character name
	
	Call Init()
	Dim characterName, questHash, param1, param2, param3, i
	charactername=Parameters(0)
	
	Dim params, ret
	If (UBound(Parameters) Mod 4 > 0) Then Call Error("Invalid number of parameters" & UBound(Parameters))
	
	For i=1 To UBound(Parameters) Step 4
		questHash=Parameters(i)
		param1=Parameters(i+1)
		param2=Parameters(i+2)
		param3=Parameters(i+3)
		
		params=Array( charactername, questHash, param1, param2, param3 )
		ret=ExecSP(characterDBconnectionString, "UpdateQuestParameter", params)
		
		If ret<>0 Then
			Call Error(ret)
		End If
	Next
	
	Call Ok(charactername)
%>
