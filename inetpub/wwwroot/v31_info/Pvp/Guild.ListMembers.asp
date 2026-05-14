<!--#include file="../common.asp"-->
<%
	' request  : guild oid
	' response : guild members (separated : character name, user type )
	' remarks  ; no error for not joined user to avoid bulky error message
	
	Call Init()
	
	Dim soap, result
	Dim guildOid
	
	guildOid	= Parameters(0)
	
	Set soap=Server.CreateObject("MSSOAP.SoapClient30")
	soap.ClientProperty("ServerHTTPRequest") = True
	Call soap.MSSoapInit(GuildWSDL)

	Dim nodeDetails, infoSet
	
	Set result=soap.GetListOfGuildMember( ServerGroupCode, guildOid )
	Set nodeDetails=result(1)
	Set infoSet=nodeDetails.childNodes(0).childNodes ' public_User_Group_GetInfo_param
	
	'Response.Write (infoSet(0).xml)
	'Response.End
	
	Dim retString, i, characterName, userType
	For i=0 To InfoSet.Length-2 ' last public_Group_User_GetList element contains recordset data
		'characterName = infoSet(i).selectSingleNode("//strNameInGroup").text
		'userType = infoSet(i).selectSingleNode("//codeGroupUserType").text
		
		characterName = infoSet(0).selectNodes("//strNameInGroup")(i).text
		userType = infoSet(0).selectNodes("//codeGroupUserType")(i).text
		
		If i>0 Then retString = retString & SEPARATOR
		retString = retString & ( characterName & SEPARATOR & userType )
	Next
	
	Response.Write (retString)
%>
