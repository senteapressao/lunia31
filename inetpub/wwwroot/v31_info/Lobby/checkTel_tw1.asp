<!--#include file="../common.asp"-->
<!--#include file="../soap_connect.asp"-->
<%
	' request  : account name
	' response : account name
	Call Init()
	If UBound(parameters)<0 Then Call Error("not enough parameter") 

	' preparing request parameters
	Dim accountName
	
	accountName	= Parameters(0)

	If isnull(accountName) OR Instr(accountName,chr(0))>0 OR accountName="" Then
		Call Error(99)
	End If

	Call Ok(accountName)
	//http://l.747a.com/v4/lobby/checkTel_tw.asp?uu0321
%>

