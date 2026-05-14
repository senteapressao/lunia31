<!--#include file="../common.asp"-->
<%
	' request  : character name, source storage page index, source storage item position, target storage page index, source storage item position
	' response : source item serial, source item hash, source stacked count, source instance,
	'			 target item serial, target item hash, target stacked count, target instance
	
	Call Init()
	Dim characterName, sourcePageIndex, sourcePosition, targetPageIndex, targetPosition
	charactername=Parameters(0)
	sourcePageIndex=Parameters(1)
	sourcePosition=Parameters(2)
	targetPageIndex=Parameters(3)
	targetPosition=Parameters(4)
	
	Dim params, retString, ret
	params=Array(charactername, sourcePageIndex, sourcePosition, targetPageIndex, targetPosition)	
	ret=ExecSP(characterDBconnectionString, "MoveCashItemInStorage", params)
	
	if ret = 0 then
		retString = params(0) & SEPARATOR & params(1) & SEPARATOR & params(2) & SEPARATOR & params(3) & SEPARATOR & params(4) & SEPARATOR & params(5) & SEPARATOR & params(6)  & SEPARATOR & params(7)
		Call Ok(retString)
	end if
	
	Call Error(ret)
%>
