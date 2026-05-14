<!--#include file="../common.asp"-->
<%
	' request  : character name[, guild oid]
	' response : guild information (guild oid, guild id, created date, guild name, intro message, master charactername, number of members)
	' remarks  ; no error for not joined user to avoid bulky error message
	
	Call Init()
	
	Dim soap, result
	Dim accountName, characterName, guildOid
	
	Dim nodeDetails, infoSet, info ' for dataset processing
	
	characterName	= Parameters(0)
	accountName=GetAccountName(characterName)
	
	If IsNull(accountName) Then Call Error(1)

	Set soap=Server.CreateObject("MSSOAP.SoapClient30")
	soap.ClientProperty("ServerHTTPRequest") = True
	Call soap.MSSoapInit(GuildWSDL)

	Set result=Nothing
	On Error Resume Next
	
	Dim retString, node
	Dim guildId, createDate, guildName, intro, master, numbers

	If UBound(Parameters)>0 Then
		guildOid=Parameters(1)
		Set result=soap.GetListOfJoinedGroup3( ServerGroupCode, guildOid, accountName, characterName )
		
		If result Is Nothing Then Call Ok("")
		
		Set nodeDetails=result(1)
		Set infoSet=nodeDetails.childNodes(0).childNodes(0) ' public_User_Group_GetInfo_param
		
		If infoSet.selectSingleNode("//oidUser_group") Is Nothing Then
			'Call Error(1) ' see the remarks
			Call Ok("") ' not joined user
		Else
		
			Set node=infoSet.selectSingleNode("//oidUser_group") : If node Is Nothing Then guildOid=0 Else guildOid=node.Text

			Set result=soap.GetGuildInfo( ServerGroupCode, guildOid )
			Set nodeDetails=result(1)
			Set infoSet=nodeDetails.childNodes(0).childNodes(0) ' public_User_Group_GetInfo_param
					
			Set node=infoSet.selectSingleNode("//strLocalID") : If node Is Nothing Then guildId="" Else guildId=node.Text
			Set node=infoSet.selectSingleNode("//dateCreate") : If node Is Nothing Then createDate="" Else createDate=node.Text
			Set node=infoSet.selectSingleNode("//strName") : If node Is Nothing Then guildName="" Else guildName=node.Text
			Set node=infoSet.selectSingleNode("//strIntro") : If node Is Nothing Then intro="" Else intro=node.Text
			Set node=infoSet.selectSingleNode("//strNameInGroup_master") : If node Is Nothing Then master="" Else master=node.Text
			Set node=infoSet.selectSingleNode("//n4RealUserCount") : If node Is Nothing Then numbers=0 Else numbers=node.Text

			retString = retString & guildOid & SEPARATOR & guildId & SEPARATOR & createDate & SEPARATOR & guildName & _
				SEPARATOR & intro & SEPARATOR & master & SEPARATOR & numbers
				
			Response.Write retString
			Response.End
		End If
	Else
		Set result=soap.GetListOfJoinedGroup( ServerGroupCode, accountName, characterName )
		If result Is Nothing Then Call Ok("")
		
		Set nodeDetails=result(1)
		Set infoSet=nodeDetails.childNodes(0).childNodes(0) ' public_User_Group_GetInfo_param
		
		If infoSet.selectSingleNode("//strLocalID_group") Is Nothing Then
			'Call Error(1) ' see the remarks
			Call Ok("") ' not joined user
		Else
			Set node=infoSet.selectSingleNode("//oidUser_group") : If node Is Nothing Then guildOid=0 Else guildOid=node.Text
			Set node=infoSet.selectSingleNode("//strLocalID_group") : If node Is Nothing Then guildId="" Else guildId=node.Text
			Set node=infoSet.selectSingleNode("//dateAccepted") : If node Is Nothing Then createDate="" Else createDate=node.Text
			Set node=infoSet.selectSingleNode("//strName_group") : If node Is Nothing Then guildName="" Else guildName=node.Text
			Set node=infoSet.selectSingleNode("//strIntro_group") : If node Is Nothing Then intro="" Else intro=node.Text
			Set node=infoSet.selectSingleNode("//strNameInGroup_master") : If node Is Nothing Then master="" Else master=node.Text
			Set node=infoSet.selectSingleNode("//n4RealUserCount") : If node Is Nothing Then numbers=0 Else numbers=node.Text

			retString = retString & guildOid & SEPARATOR & guildId & SEPARATOR & createDate & SEPARATOR & guildName & _
				SEPARATOR & intro & SEPARATOR & master & SEPARATOR & numbers
				
			Response.Write retString
			Response.End
		End If
	End If
%>
