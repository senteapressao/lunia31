<!--#include file="../common.asp"-->
<%
	' request  : character name, stage group hash, stage level
	' response : quest hashes
	
	Call Init()
	Dim characterName, stageGroupHash, stageLevel
	charactername=Parameters(0)
	stageGroupHash=Parameters(1)
	stageLevel=Parameters(2)
	
	Dim sql, connection, rs, retString, i
	Set connection=Server.CreateObject("ADODB.Connection")	
	sql = "exec ListCompletedQuests " & SqlQuot(characterName) & ", " & stageGroupHash & ", " & stageLevel
	connection.Open(characterDBconnectionString)
	Set rs=connection.Execute(sql)

	Do While Not rs.EOF
		retString = retString & rs(0) & SEPARATOR & rs(1)
		rs.MoveNext
		If Not rs.EOF Then retString = retString & SEPARATOR
	Loop
	
	rs.Close()
	connection.Close()
	Set rs=Nothing
	Set connection=Nothing
	
	' return string
	Call Ok(retString)
%>