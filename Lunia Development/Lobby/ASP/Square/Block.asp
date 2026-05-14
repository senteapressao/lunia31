<!--#include file="../common.asp"-->
<%
	' request  : character name, comment(log), [blocking time in seconds]
	' response : n/a
	
	Call Init()
	Dim characterName, comment, timeout
	characterName=Parameters(0)
	comment=Parameters(1)
	If UBound(Parameters)>1 Then timeout=CLng(Parameters(2)) Else timeout=60*60*24 ' 1 day in seconds
	
	Function ExtractQuestHash(c)
		Dim pos1, tmp
		pos1=InStr(1, c, "invalid quest(")
		if pos1>0 Then 
			tmp=Mid(c, pos1+14, InStrRev(c, ")")-(pos1+14))
			ExtractQuestHash=CLng(tmp)
		Else ' string not found
			ExtractQuestHash=0
		End If
	End Function
	
	' temporary comment filtering
	Select Case ExtractQuestHash(comment)
	Case 22742664, 54337833, 55860975, 24589708, 63292048, 182921, 21833225, 23666187, 28283791, 55261355, 55261356, 48342052, 27211772, 26288258:
		timeout=0
	End Select
	
	
	Dim params, ret
	params=Array( timeout, REQUESTERID & "(" & remoteIp & ")banned:" & SqlQuot(comment), NULL, characterName )
	ret=ExecSP(characterDBconnectionString, "Block", params)
	
	Call Ok(ret)
%>
