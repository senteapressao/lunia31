<!--#include file="../common.asp"-->
<%
	' request  : character name
	' response : quests(separated : quest hash, state, param1, param2, param3)
	
	Call Init()
	Dim characterName
	charactername=Parameters(0)
	
	Dim sql, connection, rs, retString, i
	Set connection=Server.CreateObject("ADODB.Connection")	
	sql = "exec ListWorkingQuests " & SqlQuot(characterName)
	connection.Open(characterDBconnectionString)
	Set rs=connection.Execute(sql)
	
	Do While Not rs.EOF
		retString = retString & ( rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & rs(3) & SEPARATOR & rs(4) & SEPARATOR & FormatDt(rs(5),"SQL_TM") )
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