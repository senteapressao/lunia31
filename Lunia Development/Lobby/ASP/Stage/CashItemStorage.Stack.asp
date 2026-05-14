<!--#include file="../common.asp"-->
<%
	' request  : character name, source item serial, source item hash, source item StackCount, source item Instance, Add Count,
	'			 target storage page index, target storage item position, item Stackable Count
	' response : 
	
	Call Init()
	Dim characterName, sourceSerial, sourceHash, sourceStackCount, SourceInstance, addCount, targetPageIndex, targetPosition, stackableCount
	charactername=Parameters(0)
	sourceSerial=Parameters(1)
	sourceHash=Parameters(2)
	sourceStackCount=Parameters(3)
	SourceInstance=Parameters(4)
	addCount=Parameters(5)
	targetPageIndex=Parameters(6)
	targetPosition=Parameters(7)
	stackableCount=Parameters(8)
	
	Dim params, retString, ret
	params=Array(characterName, sourceSerial, sourceHash, sourceStackCount, SourceInstance, addCount, targetPageIndex, targetPosition, stackableCount)
	ret=ExecSP(characterDBconnectionString, "StackCashItem", params)
	
	if ret = 0 then
		retString = params(0) & SEPARATOR & params(1) & SEPARATOR & params(2) & SEPARATOR & params(3) & SEPARATOR & params(4) & SEPARATOR & params(5) & SEPARATOR & params(6)
		Call Ok(retString)
	end if
	
	Call Error(ret)
%>
