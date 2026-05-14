<%	Option Explicit %>
<!--#include file="./shop_v2/inc/shop_func.asp"-->
<!--#include virtual="/lib/include/error.asp"--><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"><html xmlns="http://www.w3.org/1999/xhtml" ><head><title></title><meta http-equiv="Content-Type" content="text/html; charset=utf-8" /></head><%
	'Response.Redirect "/shop_v2/temp.asp"
	
	Dim i,j,k
	Dim plainText,encText
	Dim strAccountName,strPassword,nServerCode,strCharacterName,nLevel,nClass,oidProduct,Cat1,Cat2
	Dim currenttime,userip
	Dim codePage
	Dim retUrl

	encText = Request.QueryString
	
	plainText = Decrypt(encText)
	'println encText
	'println plainText
	On Error Resume Next
	strAccountName	= Split(Split(plainText,"&")(0),"=")(1)
	strPassword		= Split(Split(plainText,"&")(1),"=")(1)
	nServerCode		= Split(Split(plainText,"&")(2),"=")(1)
	strCharacterName= Split(Split(plainText,"&")(3),"=")(1)
	nLevel			= Split(Split(plainText,"&")(4),"=")(1)
	nClass			= Split(Split(plainText,"&")(5),"=")(1)
	
	currenttime		= Split(Split(plainText,"&")(6),"=")(1)
	userip			= Split(Split(plainText,"&")(7),"=")(1)
	If CountryCode="EU" Then
		codePage	= Split(Split(plainText,"&")(8),"=")(1)
		oidProduct	= Split(Split(plainText,"&")(9),"=")(1)
		Cat1		= Split(Split(plainText,"&")(10),"=")(1)
		Cat2		= Split(Split(plainText,"&")(11),"=")(1)
	ElseIf CountryCode="MY" Then
		codePage	= Split(Split(plainText,"&")(8),"=")(1)
		oidProduct	= Split(Split(plainText,"&")(9),"=")(1)
		Cat1		= Split(Split(plainText,"&")(10),"=")(1)
		Cat2		= Split(Split(plainText,"&")(11),"=")(1)
	Else
		oidProduct	= Split(Split(plainText,"&")(8),"=")(1)
		Cat1		= Split(Split(plainText,"&")(9),"=")(1)
		Cat2		= Split(Split(plainText,"&")(10),"=")(1)
	End If
	
	'println strAccountName	
	'println strPassword		
	'println nServerCode		
	'println strCharacterName
	'println nClass
	'println codePage
	'println "codePage"
	'Response.End

	If CountryCode="EU" or CountryCode="MY" Then
		If codePage = "French" Then
			codePage = "fr"
		ElseIf codePage = "MY_Chinese" Then
			codePage = "zh-cn"
		Else
			codePage = ""
		End If
		Response.Cookies("UserSession")("codePage") = codePage
	End If
	
	If Err.number>0 Then
	End If
	On Error Goto 0
	
	'If DateDiff("n",currenttime,Now())<-1 OR DateDiff("n",currenttime,Now())>1 Then Call ThrowError(null,"illegal access:5","close")
	'If userip<>Request.ServerVariables("REMOTE_ADDR") Then Call ThrowError(null,"illegal access:6","close")
	If strAccountName="" Then Call ThrowError(null,"illegal access:7","close")
	
	If oidProduct<>"" Then retUrl="/shop_v2/read.asp?oidProduct="& oidProduct
	If Cat1<>"" Then retUrl="/shop_v2/list.asp?Cat1="& Cat1 &"&Cat2="& Cat2
	If retUrl="" Then retUrl="/shop_v2/list.asp"

	Function Encrypt(plainText)
		Dim oEnc,password
		password = "Wonder Girls"
		set oEnc = Server.CreateObject("DynuEncrypt.Functions")		
		
		Encrypt = oEnc.Encrypt(plainText,password)
		
		set oEnc = Nothing		
	End Function
	
	Function Decrypt(encText)
		Dim oEnc,password
		password = "Wonder Girls"
		set oEnc = Server.CreateObject("DynuEncrypt.Functions")		
		
		Decrypt = oEnc.Decrypt(encText,password)
		
		set oEnc = Nothing		
	End Function
%>
<body  style="margin:0px;" onload="frm_login.submit();"><form id="frm_login" name="frm_login" method="post" action="/shop_v2/login/login_post.asp" onsubmit="validate_login();return false;"><input type="hidden" name="RetURL" value="<%=retUrl %>" /><input type="hidden" name="strAccountName" value="<%=strAccountName %>" /><input type="hidden" name="strPassword" value="<%=strPassword %>" /><input type="hidden" name="nServerCode" value="<%=nServerCode %>" /><input type="hidden" name="strCharacterName" value="<%=strCharacterName %>" /><input type="hidden" name="nLevel" value="<%=nLevel %>" /><input type="hidden" name="nClass" value="<%=nClass %>" /></form></body></html>