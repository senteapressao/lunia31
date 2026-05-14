<%	Option Explicit %>
<!--#include file="./common.asp"-->
<%
	Dim i,j,k
	Dim arrCat
	Dim retString : retString = ""
	
	arrCat = Array( _
		Array(1,0,"騎士",1) _
	,		Array(1,1,"裝備",1)	 _
	,	Array(2,0,"治癒師",0) _
	,		Array(2,1,"裝備",0) _
	,	Array(3,0,"魔法師",2) _
	,		Array(3,1,"裝備",2) _
	,	Array(4,0,"操偶師",5) _
	,		Array(4,1,"裝備",5) _
	,	Array(5,0,"盜賊",3) _
	,		Array(5,1,"裝備",3) _
	,	Array(6,0,"史萊姆",4) _
	,		Array(6,1,"裝備",4) _
	,	Array(7,0,"克里克",6) _
	,		Array(7,1,"裝備",6) _
	,	Array(8,0,"雪女",7) _
	,		Array(8,1,"裝備",7) _
	,	Array(9,0,"弓箭手",8) _
	,		Array(9,1,"裝備",8) _
	,	Array(10,0,"黎恩",9) _
	,		Array(10,1,"裝備",9) _
	,	Array(11,0,"卡莉",10) _
	,		Array(11,1,"裝備",10) _
	,	Array(12,0,"明日香",11) _
	,		Array(12,1,"裝備",11) _
	,	Array(98,0,"寵物",10000) _
	,		Array(98,1,"寵物",10000) _
	,		Array(98,2,"裝備",10000) _
	,	Array(99,0,"福袋",9999) _
	,		Array(99,1,"福袋",9999) _
	)
	
	For i=0 To Ubound(arrCat)
		If IsArray(arrCat(i)) Then
			retString = retString & SEPARATOR
			retString = retString & arrCat(i)(0) &_
				SUBSEPARATOR & arrCat(i)(1) &_
				SUBSEPARATOR & arrCat(i)(2) &_
				SUBSEPARATOR & arrCat(i)(3)
		End If
	Next
	
	Call OK(retString)
%>