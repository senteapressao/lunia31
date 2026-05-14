<!--#include file="../common.asp"-->
<%
	' request  : character name, quest hash, quest stage group hash, quest stage level, next quest stage group hash, next quest stage level, select item index
	' response : quest hash, select item index
	
	Call Init()
	Dim characterName, questHash, stageGroupHash, stageLevel, nextQuestStageGroupHash, nextQuestStageLevel, selectItemIndex,deleteWorkingQuest
	charactername=Parameters(0)
	questHash=Parameters(1)
	stageGroupHash=Parameters(2)
	stageLevel=Parameters(3)
	nextQuestStageGroupHash=Parameters(4)
	nextQuestStageLevel=Parameters(5)
	deleteWorkingQuest=Parameters(6)
	If UBound(Parameters)>6 Then selectItemIndex=Parameters(7)
	
	Dim params, ret
	params=Array( charactername, questHash, stageGroupHash, stageLevel, nextQuestStageGroupHash, nextQuestStageLevel, deleteWorkingQuest)
	ret=ExecSP(characterDBconnectionString, "CompleteQuest", params)
	
	If ret=0 Then
		Call Ok(questHash & SEPARATOR & selectItemIndex)
	Else
		Call Error(ret)
	End If
%>
