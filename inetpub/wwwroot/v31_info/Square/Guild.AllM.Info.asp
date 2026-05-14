<!--#include file="../common.asp"-->
<%
	' request  : characterName
	' response : guild Info (separated : guildId, guildName, message, masterName, memberCount, myGrade, myGuildMemberId, guildLevel, [gradeName, authority, gradeName, authority...] )

	Call Init()

	Dim connection, command, retString, rs

	Set connection=Server.CreateObject("ADODB.Connection")
	Set command=Server.CreateObject("ADODB.Command")			

	connection.Open(guildDBconnectionString)
	command.ActiveConnection=connection
	command.CommandType = 4 ' adCmdStoredProc
	command.CommandText = "GuildInfo"
	command.Parameters(1)= Parameters(0)

	Set rs = command.Execute()
	retString = ""
	Do While Not rs.EOF ' Parsing guild info and my grade
		retString = retString & ( rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & rs(3) & SEPARATOR & rs(4) & SEPARATOR & rs(5) & SEPARATOR & rs(6) & SEPARATOR & rs(7) )
		rs.MoveNext
		If Not rs.EOF Then retString = retString & SEPARATOR
	Loop

	retString = retString & SEPARATOR
	
	Set rs= rs.NextRecordset ' Parsing guild grades
	Do While Not rs.EOF
		retString = retString & ( rs(0) & SUBSEPARATOR & rs(1) )
		rs.MoveNext
		If Not rs.EOF Then retString = retString & SUBSEPARATOR
	Loop

	rs.Close()
	connection.Close()

	Set command=Nothing
	Set rs = Nothing
	Set connection=Nothing

	' return string
	Call Ok(retString)
%>
