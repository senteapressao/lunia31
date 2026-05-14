<!--#include file="../common.asp"-->
<%
	' request  : characterName, petSerial, itemHash, itemStackCount, itemInstance, expFactor, trainingStartTime, trainingFinishTime
	' response : petSerial, itemHash, itemStackCount, itemInstance, expFactor, trainingStartTime, trainingFinishTime
	
	Call Init()
	
	dim characterName, petSerial, itemHash, itemStackCount, itemInstance, expFactor, trainingStartTime, trainingFinishTime
	
	characterName=Parameters(0)
	petSerial=Parameters(1)
	itemHash=Parameters(2)
	itemStackCount=Parameters(3)
	itemInstance=Parameters(4)
	expFactor=Parameters(5)
	trainingStartTime=Parameters(6)
	trainingFinishTime=Parameters(7)
	
	Dim params, retString, ret
	params=Array(characterName, petSerial, itemHash, itemStackCount, itemInstance, expFactor, trainingStartTime, trainingFinishTime)
	ret=ExecSP(characterDBconnectionString, "PetStartTraining", params)
		
	if ret = 0 then
		retString = params(0) & SEPARATOR & params(1) & SEPARATOR & params(2) & SEPARATOR & params(3) & SEPARATOR & params(4) & SEPARATOR &  FormatDatetime(params(5), 2) &" "& FormatDatetime(params(5), 4) & SEPARATOR &  FormatDatetime(params(6), 2) &" "& FormatDatetime(params(6), 4)
		Call Ok(retString)
	end if
	
	Call Error(ret)
%>
