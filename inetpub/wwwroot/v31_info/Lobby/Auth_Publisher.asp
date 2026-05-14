<!--#include file="../common.asp"-->
<%
	' request  : account name, password
	' response : account name
	Call Init()
	If UBound(parameters)<1 Then Call Error("not enough parameter") 

	' preparing request parameters
	Dim accountName, password, ipadder
	
	accountName=Parameters(0)
	password=Parameters(1)
	ipadder =Parameters(2)
	' check parameter validation

	' call Stored procedure	
	Dim params, ret
	params=Array(REQUESTERID, accountName, password,ipadder)
	'Response.Write ipadder
	ret=ExecSP(characterDBconnectionString, "AuthAccount", params)
	If ret=0 Then
		Call Ok(accountName)
	Else
		Call Error(ret)
	End If
%>

