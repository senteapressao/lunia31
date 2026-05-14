<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/helpers.asp"-->
<!--#include virtual="/include/common.asp"-->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	'Parameters = Split("0Teles9999Leandro01", SEPARATOR)
	Dim stgType, playerName, classSearch, majorStatus, playerSearch, nRankSearch,pageNumber
	If UBound(parameters)<5 Then Call Error("not enough parameter")
	playerName	= Parameters(1)
	classSearch = Parameters(2)
	playerSearch = Parameters(3)
	nRankSearch = CInt(Parameters(4))
	pageNumber = CInt(Parameters(5))
	if pageNumber = 0 then 'In case is the first search lunia sends in pageNumber = 0
		pageNumber = 1
	end if
	
	dim filename
	filename = getCacheFile()
	
	
	dim fs,f
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	set f=fs.GetFile(filename)
	if DateAdd("h",24,CDate(f.DateLastModified)) < Now() then
		UpdateAchievement(filename)
	end if
	set f=nothing
	set fs=nothing
	
	Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument.3.0")    
	objXMLDoc.async = False    
	objXMLDoc.load filename
	
	
	dim TotalPageNumber : TotalPageNumber = 1
	'if no searchs are available we do the math to get the rank total pages.
	'doom,,,,richenessforest,MythTower,summonscave,doomtower,doomgateway,exercise
	dim i : i = 1 'Position on the ranking
	dim j : j = 1 'Total player per page counter
	dim InsertedList
	redim InsertedList(0)
	dim positionNumberR : positionNumberR = "-"
	dim playerNameR : playerNameR = "-"
	dim classNumberR  : classNumberR = 0
	dim achievementPointR : achievementPointR = 0
	dim retString : retString = ""
	
	for each item in objXMLDoc.selectNodes("/rank/achievement/player")
		'dim positionNumber : positionNumber = item.selectSingleNode("positionNumber").text 
		dim characterName : characterName 		= item.selectSingleNode("characterName").text
		dim classNumber : classNumber 			= item.selectSingleNode("classNumber").text 
		dim achievementPoint : achievementPoint = item.selectSingleNode("achievementPoint").text 
		
		if classSearch = "9999" or classSearch = classNumber then
			if characterName = playerName then
				positionNumberR = i
				playerNameR = characterName
				classNumberR = classNumber
				achievementPointR = achievementPoint
			end if
			
			if nRankSearch > 0 or playerSearch <> "" then
				if nRankSearch = i or InStr(LCase(characterName),LCase(playerSearch)) > 0 then
					retString = retString &_
						SEPARATOR 	 & i &_
						SUBSEPARATOR & characterName &_
						SUBSEPARATOR & classNumber &_
						SUBSEPARATOR & achievementPoint
					j = j + 1 ' items in the page
					if nRankSearch = i then Exit For end if
				end if
			elseif roundUp(i/ranksPerPage) = pageNumber then
				'response.write "Position=" & i & "characterName=" & characterName & "|Stats=" &status_Amt & "<br>"
				retString = retString &_
					SEPARATOR 	 & i &_
					SUBSEPARATOR & characterName &_
					SUBSEPARATOR & classNumber &_
					SUBSEPARATOR & achievementPoint
				j = j + 1 ' items in the page
			end if
			i = i + 1
		end if
	Next
	'response.write i
	if playerSearch = "" and nRankSearch = 0 then
		TotalPageNumber = roundUp(i/ranksPerPage)
	end if
	myData = positionNumberR & SUBSEPARATOR & playerNameR & SUBSEPARATOR & classNumberR  & SUBSEPARATOR & achievementPointR
	' @Teles :)
	Call Ok(TotalPageNumber & SEPARATOR & pageNumber & SEPARATOR & myData & retString)
%>

<%
function CacheMajor(RS, majorType, objFile)
	objFile.WriteLine("<"&majorType&">")
	i = 0
	Do Until RS.Eof
		i=i+1
		objFile.WriteLine("	<player>" )
		objFile.WriteLine("		<positionNumber>" & i & "</positionNumber>" )
		objFile.WriteLine("		<characterName>" & rs(0) & "</characterName>" )
		objFile.WriteLine("		<classNumber>" & rs(1) & "</classNumber>" )
		objFile.WriteLine("		<achievementPoint>" & rs(2) & "</achievementPoint>" )
		objFile.WriteLine("	</player>" )
		rs.MoveNext
	Loop
	objFile.WriteLine("</"&majorType&">")
end function
function UpdateAchievement(fileName)
	' fileName for the cache file.
	dim fs,destinationPath
	destinationPath = fileName
	set fs = CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(destinationPath) then
		fs.DeleteFile(destinationPath)
	end if
	set objFile = fs.OpenTextFile(destinationPath, 2, true, -2) 'Opening the file and storing it's pointer to variable
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = false
		Set .cmd = Command
		.ConnStr = ConnStrCharacter
		.SPName = "dbo.Rank_ListAchievement"
		Call .InitCommand()
		blsResult = .ExecRecordset()
	End with

	If blsResult=False Then		
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
	objFile.WriteLine("<rank>")
		Set RS = sph.rs : If TypeName(RS)<>"Recordset" Then Call Error("no recordset1")
		CacheMajor RS, "achievement", objFile
	objFile.WriteLine("</rank>")
	End If
	
	set sph = Nothing

	RS.Close
end function

%>