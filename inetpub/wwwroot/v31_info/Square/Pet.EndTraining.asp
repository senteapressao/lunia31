<!--#include file="../common.asp"-->
<%
	' request  : characterName, petSerial
	' response : petSerial, itemHash, itemStackCount, itemInstance, expFactor, trainingStartTime, trainingFinishTime
	
	Call Init()
	
	dim characterName, petSerial
	
	characterName=Parameters(0)
	petSerial=Parameters(1)
	
	Dim params, retString, ret
	params=Array(characterName, petSerial)
	ret=ExecSP(characterDBconnectionString, "PetEndTraining", params)
	
	if ret = 0 then
		retString = params(0) & SEPARATOR & params(1) & SEPARATOR & params(2) & SEPARATOR & params(3) & SEPARATOR & params(4) & SEPARATOR & params(5) & SEPARATOR & params(6)
		Call Ok(retString)
	end if
	
	Call Error(ret)
%>
