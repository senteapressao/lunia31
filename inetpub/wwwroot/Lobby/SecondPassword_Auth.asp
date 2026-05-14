<!--#include file="../common.asp"-->
<%
	' request  : account name, password
	' response : account name
	Call Init()
	If UBound(parameters)<1 Then Call Error("not enough parameter") 

	' preparing request parameters
	Dim accountName, password
	
	accountName=Parameters(0)
	password=Parameters(1)
	
	' check parameter validation

	' call Stored procedure	
	Dim params, ret
	params=Array(accountName, password)
	
	ret=ExecSP(characterDBconnectionString, "SecondPassword_Auth", params)
	If ret=0 Then
		Call Ok("Por que está entrando aqui em?")
	Else
		Call Error(ret)
	End If
%>