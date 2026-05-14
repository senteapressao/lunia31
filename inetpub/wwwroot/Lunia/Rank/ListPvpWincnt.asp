<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/helpers.asp"-->
<!--#include virtual="/include/common.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%  
	Parameters = Split("0Teleswk401",SEPARATOR)	
	Dim playerName, pageNumber, playerSearch, classSearch
	If UBound(parameters)<5 Then Call Ok("11--00-1")
	
	playerName	= Parameters(1)
	rankType = Parameters(2)
	teamSize = Parameters(3)
	playerSearch = Parameters(4)
	rankNSearch = Parameters(5)
	pageNumber = Parameters(6) 
	
	Dim inside, pageNumberMath
	
	Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument.3.0")    
	objXMLDoc.async = False    
	objXMLDoc.load getCacheFile()
	
	dim i : i = 1
	dim positionNumberR : positionNumberR = "-"
	dim playerNameR : playerNameR = "-"
	dim classNumberR  : classNumberR = 0
	dim floorR : floorR = 0
	dim playTimeR : playTimeR = 0
	
	dim retString : retString = ""
	for each item in objXMLDoc.selectNodes("/rank/"&searchParam2&"/player")
		dim positionNumber : positionNumber = item.selectSingleNode("positionNumber").text 
		dim characterName : characterName 	= item.selectSingleNode("characterName").text 
		dim classNumber : classNumber 		= item.selectSingleNode("classNumber").text 
		dim floor : floor 					= item.selectSingleNode("floor").text 
		dim playTime : playTime 			= item.selectSingleNode("playTime").text 
		dim stageGroupHash : stageGroupHash = item.selectSingleNode("stageGroupHash").text
		if searchParam1=stageGroupHash then
			if playerSearch = "" then
				' add count to classNumber
				if characterName = playerName then ' If playerName matches with the requested user
					positionNumberR = i
					playerNameR = characterName
					classNumberR = classNumber
					stageLevelR = stageLevel
					floorR = floor
					playTimeR = playTime
				end if
				if roundUp(i/20) = pageNumber or classSearch = 9999 or classSearch = classNumber then
					retString = retString &_
						SEPARATOR & i &_
						SUBSEPARATOR & characterName &_
						SUBSEPARATOR & classNumber &_
						SUBSEPARATOR & floor &_
						SUBSEPARATOR & playTime
				end if
				i = i+1
			elseif playerSearch = characterName then
				retString = retString &_
					SEPARATOR & i &_
					SUBSEPARATOR & characterName &_
					SUBSEPARATOR & classNumber &_
					SUBSEPARATOR & floor &_
					SUBSEPARATOR & playTime
					i = 1
				Exit for
			end if
		end if
	Next
	TotalPageNumber = roundUp(i/20)
	if TotalPageNumber = 0 then TotalPageNumber = 1 end if
	
	If Cint(pageNumber) > Cint(TotalPageNumber) then Call Error(Cint(pageNumber) & ">" & Cint(TotalPageNumber))
	myData = positionNumberR & SUBSEPARATOR & playerNameR & SUBSEPARATOR & classNumberR & SUBSEPARATOR & FloorR & SUBSEPARATOR & playTimeR
	Call Ok(TotalPageNumber & SEPARATOR & pageNumber & SEPARATOR & myData & retString)
%>