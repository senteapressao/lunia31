<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/helpers.asp"-->
<!--#include virtual="/include/common.asp"-->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	'Parameters = Split("01Teles9999STR01", SEPARATOR)
	'Parameters = Split("00Teles9999STR00", SEPARATOR)
	Dim stgType, playerName, classSearch, majorStatus, playerSearch, nRankSearch,pageNumber
	If UBound(parameters)<7 Then Call Error("not enough parameter")
	
	serverCode = parameters(0) 			'0		?
	stgType = CInt(parameters(1))
	playerName	= Parameters(2)			'Teles	User CharacterName
	classSearch = Parameters(3)			'9999	ClassType
	majorStatus = Parameters(4)			'STR 	MajorType
	playerSearch = LCase(Parameters(5))	'Moki 	search by playerName
	nRankSearch = CInt(Parameters(6))	'0 		search by rank
	pageNumber = CInt(parameters(7))	'0 		which page the user is at
	if pageNumber = 0 then 'In case is the first search lunia sends in pageNumber = 0
		pageNumber = 1
	end if
	
	dim filename
	filename = CStr(getCacheFile())
	
	dim fs,f
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(filename) then
		set f=fs.GetFile(filename)
		if DateAdd("h",24,CDate(f.DateLastModified)) < Now() then
			UpdateListStat(filename)
		end if
		set f=nothing
	end if
	set fs=nothing

	Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument.6.0") 
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
	dim status_AmtR : status_AmtR = 0
	dim stageLevelR : stageLevelR = 0
	dim retString : retString = ""
	dim query
	
	If  stgType = 0 Then
		set query =  objXMLDoc.selectNodes("/rank/stage/"&majorStatus&"/player")
	Else
		set query =  objXMLDoc.selectNodes("/rank/myth/"&majorStatus&"/player")
	End if

	for each item in query 
		'dim positionNumber : positionNumber = item.selectSingleNode("positionNumber").text 
		dim characterName : characterName 	= item.selectSingleNode("characterName").text
		dim classNumber : classNumber 		= item.selectSingleNode("classNumber").text 
		dim status_Amt : status_Amt 		= item.selectSingleNode("value").text
		dim stageLevel : stageLevel 		= item.selectSingleNode("stageLevel").text
		
		'response.write characterName & i & "<br>"
		if (classSearch = "9999" or classSearch = classNumber) then
			if nRankSearch > 0 or playerSearch <> "" then
				if nRankSearch = i or InStr(LCase(characterName),playerSearch) > 0 then
					retString = retString &_
						SEPARATOR 	 & i &_
						SUBSEPARATOR & i &_
						SUBSEPARATOR & characterName &_
						SUBSEPARATOR & classNumber &_
						SUBSEPARATOR & status_Amt &_
						SUBSEPARATOR & stageLevel &_
						SUBSEPARATOR & 0 &_
						SUBSEPARATOR & 0
					j = j + 1 ' items in the page
					if nRankSearch = i then Exit For end if
				end if
			elseif roundUp(i/ranksPerPage) = pageNumber then
				
				retString = retString &_
					SEPARATOR 	 & i &_
					SUBSEPARATOR & i &_
					SUBSEPARATOR & characterName &_
					SUBSEPARATOR & classNumber &_
					SUBSEPARATOR & status_Amt &_
					SUBSEPARATOR & stageLevel &_
					SUBSEPARATOR & 0 &_
					SUBSEPARATOR & 0
				j = j + 1 ' items in the page
			end if
			i = i + 1
		end if
	Next
	'response.write i
	if playerSearch = "" and nRankSearch = 0 then
		TotalPageNumber = roundUp(i/ranksPerPage)
	end if
	myData = positionNumberR & SUBSEPARATOR & positionNumberR & SUBSEPARATOR & playerNameR & SUBSEPARATOR & classNumberR & SUBSEPARATOR & status_AmtR & SUBSEPARATOR & stageLevelR & SUBSEPARATOR & 0 & SUBSEPARATOR & 0
	' @Teles :)
	Call Ok(TotalPageNumber & SEPARATOR & pageNumber & SEPARATOR & myData & retString)
%>
<%
Function UpdateListStat(fileName)
	Server.ScriptTimeout = 200
	' fileName for the cache file.
	dim fs,destinationPath
	destinationPath = fileName
	set fs = CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(destinationPath) then
		fs.DeleteFile(destinationPath)
	end if
	set objFile = fs.OpenTextFile(destinationPath, 2, true, -2) 'Opening the file and storing it's pointer to variable
	
	objFile.WriteLine("<rank>")
		' call Stored procedure
		Set sph = new SPHelper
		with sph
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = ConnStrCharacter
			.SPName = "dbo.Rank_ListStat"
			Call .InitCommand()
			blsResult = .ExecRecordset8()
		End with

		If blsResult=False Then		
			if sph.frk_n4ErrorCode>0 Then
				Call Error(sph.frk_n4ErrorCode)
			Else
				Call Error(sph.frk_strErrorText)
			End If
		Else
			objFile.WriteLine("	<stage>")

				Set rs1 = sph.rs : If TypeName(rs1)<>"Recordset" Then Call Error("no recordset1")
				CacheMajor rs1, "STR", objFile

				Set rs2 = sph.rs2 : If TypeName(rs2)<>"Recordset" Then Call Error("no recordset2")
				CacheMajor rs2, "DEX", objFile
				
				Set rs3 = sph.rs3 : If TypeName(rs3)<>"Recordset" Then Call Error("no recordset3")
				CacheMajor rs3, "VIT", objFile

				Set rs4 = sph.rs4 : If TypeName(rs4)<>"Recordset" Then Call Error("no recordset4")
				CacheMajor rs4, "INT", objFile
				

			objFile.WriteLine("	</stage>")
			objFile.WriteLine("	<myth>")

				Set rs5 = sph.rs5 : If TypeName(rs5)<>"Recordset" Then Call Error("no recordset5")
				CacheMajor rs5, "STR", objFile
				
				Set rs6 = sph.rs6 : If TypeName(rs6)<>"Recordset" Then Call Error("no recordset6")
				CacheMajor rs6, "DEX", objFile
				
				Set rs7 = sph.rs7 : If TypeName(rs7)<>"Recordset" Then Call Error("no recordset7")
				CacheMajor rs7, "VIT", objFile
				
				Set rs8 = sph.rs8 : If TypeName(rs8)<>"Recordset" Then Call Error("no recordset8")
				CacheMajor rs8, "INT", objFile

			objFile.WriteLine("	</myth>")

			rs1.Close
			rs2.Close
			rs3.Close
			rs4.Close
			rs5.Close
			rs6.Close
			rs7.Close
			rs8.Close
		End If
		set sph = nothing
	objFile.WriteLine("</rank>")
	

	
End Function

Function CacheMajor(rs, majorType, objFile)
	objFile.WriteLine("		<"&majorType&">")
		Do Until rs.Eof
			objFile.WriteLine("			<player>" )
			objFile.WriteLine("				<characterName>" & rs(0) & "</characterName>" )
			objFile.WriteLine("				<classNumber>" & rs(1) & "</classNumber>" )
			objFile.WriteLine("				<stageLevel>" & rs(2) & "</stageLevel>" )
			objFile.WriteLine("				<value>" & rs(3) & "</value>" )
			objFile.WriteLine("			</player>" )
			rs.MoveNext
		Loop
	objFile.WriteLine("		</"&majorType&">")
End Function
%>