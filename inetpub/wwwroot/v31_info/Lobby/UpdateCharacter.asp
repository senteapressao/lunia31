<!--#include file="../common.asp"-->
<%
	' request  : character name
	' response : character name, class number, level, licenses(subseparated), equipments(subseparated: item hash, position, instance), character licenses(subseparated)
	 
	Call Init()
	If UBound(parameters)<0 Then Call Error("not enough parameter")

	' preparing request parameters
	Dim characterName
	characterName=Parameters(0)
	
	Dim sql, connection, rs, retString, i
	Set connection=Server.CreateObject("ADODB.Connection")	

	' sql query to retrieve basic information of characters
	sql="Exec InfoCharacterForLobby " & SqlQuot(characterName)
	
	Dim CharacterBaseString, CharacterBases
	
	connection.Open(characterDBconnectionString)
	Set rs=connection.Execute(sql)
	CharacterBaseString=rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & _
		rs(3) & SEPARATOR & rs(4) & SEPARATOR & rs(5) & SEPARATOR & rs(6) & SEPARATOR

	' append license information
	Set rs= rs.NextRecordset
	Do While Not rs.EOF
		CharacterBaseString = CharacterBaseString & rs(0) & SUBSEPARATOR & rs(1)
		rs.MoveNext
		If Not rs.EOF Then CharacterBaseString = CharacterBaseString & SUBSEPARATOR
	Loop
	CharacterBaseString = CharacterBaseString & SEPARATOR
		
	' append equipment information
	Set rs= rs.NextRecordset 
	Do While Not rs.EOF
		CharacterBaseString = CharacterBaseString & ( rs(0) & SUBSEPARATOR & rs(1) & SUBSEPARATOR & rs(2) )
		rs.MoveNext
		If Not rs.EOF Then CharacterBaseString = CharacterBaseString & SUBSEPARATOR
	Loop
	CharacterBaseString = CharacterBaseString & SEPARATOR
	
	' append character licenses
	Set rs= rs.NextRecordset 
	Do While Not rs.EOF
		CharacterBaseString = CharacterBaseString & rs(0)
		rs.MoveNext
		If Not rs.EOF Then CharacterBaseString = CharacterBaseString & SUBSEPARATOR
	Loop
	
	rs.Close
	Set rs=Nothing
	connection.Close()
	Set connection=Nothing
	
	' return the result string
	Call Ok(CharacterBaseString)
%>
