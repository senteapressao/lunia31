<!--#include file="../common.asp"-->
<%
	' request  : character name, source storage page index, source storage item position
	' response : item serial, item hash, stacked count, instance, source storage page index, source storage item position
	
	Call Init()
	Dim characterName, pageIndex, position
	charactername=Parameters(0)
	pageIndex=Parameters(1)
	position=Parameters(2)
	
	Dim sql, connection, rs, retString, i
	Set connection=Server.CreateObject("ADODB.Connection")	
	sql = "exec PopCashItem " & SqlQuot(characterName) & ", " & pageIndex & ", " & position
	
	connection.Open(characterDBconnectionString)
	Set rs = connection.Execute( sql )
	
	if rs.EOF Then
		Call Error(1) ' invalid item serial
	Else
		retString = retString & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & rs(3)
	End If

	connection.Close()
	Set rs=Nothing
	Set connection=Nothing
	
	retString = retString & SEPARATOR & pageIndex & SEPARATOR & position
	
	' return string
	Call Ok(retString)
%>
