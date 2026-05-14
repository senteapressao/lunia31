<%	Option Explicit %>
<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/common.asp"-->
<%
	Dim i,j,k
	Dim arrCat
	Dim retString : retString = ""
	
	Sql = "select stRSlot,oidProduct,itemHash from dbo.tblSlot (nolock)"
	
	Conn.Open(ConnStrShop)
	Set RS = conn.Execute(SQL)
	
	Do Until RS.Eof	
		If retString<>"" Then retString = retString & SEPARATOR
		retString = retString & RS(0) &_
			SUBSEPARATOR & RS(1) &_
			SUBSEPARATOR & RS(2)
	
		RS.MoveNext
	Loop
	RS.Close
	Conn.Close
	
	Call Ok(retString)
%>