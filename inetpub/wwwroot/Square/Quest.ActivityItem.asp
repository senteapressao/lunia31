<!--#include file="../common.asp"-->
<%
	' request  : character name, quest hashes
	' response : quest hashes
	
	Call Init()
	
	If UBound(Parameters)>1 Then ' item exists
		
		Dim characterName
		charactername=Parameters(0)
		
		Dim sql, connection, rs, retString, i
		Set connection=Server.CreateObject("ADODB.Connection")	
		connection.Open(characterDBconnectionString)

		for i=1 To UBOUND( Parameters )
			
			sql = "exec CompletedActiveItemQuest " & SqlQuot(characterName) & ", " & Parameters(i)
			Set rs=connection.Execute(sql)
			
			Do While Not rs.EOF
				retString = retString & rs(0)
				If Not i=UBOUND( Parameters ) Then retString = retString & SEPARATOR
				rs.MoveNext
				'If Not rs.EOF Then retString = retString & SEPARATOR
			Loop

		NEXT
		
		rs.Close()
		connection.Close()
		Set rs=Nothing
		Set connection=Nothing
	
	End If
	
	' return string
	Call Ok(retString)
%>