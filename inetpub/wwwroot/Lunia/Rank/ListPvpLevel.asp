<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/helpers.asp"-->
<!--#include virtual="/include/common.asp"-->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	'Parameters = Split("0Teles999901", SEPARATOR)
	' preparing request parameters
	Dim playerName, pageNumber, playerSearch, classSearch
	If UBound(parameters)<5 Then Call Error("not enough parameter")
	playerName	= Parameters(1)
	classSearch = Parameters(2)
	playerSearch = Parameters(3)
	pageNumber = Parameters(5)
	
	dim filename
	filename = getCacheFile()
	
	
	dim fs,f
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	set f=fs.GetFile(filename)
	if DateAdd("h",24,CDate(f.DateLastModified)) < Now() then
		UpdatePvpLevel(filename)
	end if
	set f=nothing
	set fs=nothing

	Dim inside, pageNumberMath
	
	Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument.3.0")    
	objXMLDoc.async = False 
	objXMLDoc.load(filename)
	
	Dim xmlProduct
	dim i : i = 0
	dim playerNameR : playerNameR = "-"
	dim classNumberR  : classNumberR = "0"
	dim pvpLevelR : pvpLevelR = "-"
	dim pvpExpR : pvpExpR = "-"
	dim positionNumberR : positionNumberR = "-"
	
	dim TotalPageNumber
	TotalPageNumber = roundUp(objXMLDoc.documentElement.selectNodes("player").length/20)
	If Cint(pageNumber) > Cint(TotalPageNumber) then Call Error(Cint(pageNumber) & ">" & Cint(TotalPageNumber))
	
	if pageNumber > 1 then pageNumberMath = pageNumber - 1
	dim classSearchRank : classSearchRank = 1
	For Each xmlProduct In objXMLDoc.documentElement.selectNodes("player")
		i=i+1
		Dim positionNumber : positionNumber = xmlProduct.selectSingleNode("positionNumber").text
		Dim characterName : characterName = xmlProduct.selectSingleNode("characterName").text
		Dim classNumber : classNumber = xmlProduct.selectSingleNode("classNumber").text
		Dim pvpLevel : pvpLevel = xmlProduct.selectSingleNode("pvpLevel").text
		Dim pvpExp : pvpExp = xmlProduct.selectSingleNode("pvpExp").text
		if characterName = playerName then
			playerNameR = playerName
			pvpLevelR = pvpLevel
			classNumberR = classNumber
			pvpExpR = pvpExp
			positionNumberR = positionNumber
		end if
			if playerSearch <> "" and characterName = playerSearch then
				pageNumber=roundUp(positionNumber/20)
				pageNumberMath=roundUp(positionNumber/20)-1
				playerSearch=""
			end if
			if classSearch <> "9999" then 
				if playerSearch = "" then
					if (Server.HTMLEncode(classNumber) = classSearch) and classSearchRank < 21 then
						inside = inside & SEPARATOR & Cstr(classSearchRank) & SUBSEPARATOR & Cstr(classSearchRank)
						inside = inside & SUBSEPARATOR & Server.HTMLEncode(characterName)
						inside = inside & SUBSEPARATOR & Server.HTMLEncode(classNumber)
						inside = inside & SUBSEPARATOR & Server.HTMLEncode(pvpLevel)
						inside = inside & SUBSEPARATOR & Server.HTMLEncode(pvpExp)
						classSearchRank = classSearchRank & 1
					end if
				end if
			else
				If i>(0+(20*pageNumberMath)) and i <(21+(20*pageNumberMath)) then
					if playerSearch = "" then
						inside = inside & SEPARATOR & Server.HTMLEncode(positionNumber) & SUBSEPARATOR & Server.HTMLEncode(positionNumber)
						inside = inside & SUBSEPARATOR & Server.HTMLEncode(characterName)
						inside = inside & SUBSEPARATOR & Server.HTMLEncode(classNumber)
						inside = inside & SUBSEPARATOR & Server.HTMLEncode(pvpLevel)
						inside = inside & SUBSEPARATOR & Server.HTMLEncode(pvpExp)
					end if
				end If
			end if
				
	Next
		if classSearch <> "9999" then
			TotalPageNumber = 1
		end if
	If Cint(pageNumber) > Cint(TotalPageNumber) then Call Error(Cint(pageNumber) & ">" & Cint(TotalPageNumber))
	str = Cstr(TotalPageNumber) & SEPARATOR & Cstr(pageNumber) & SEPARATOR & "1" & SUBSEPARATOR & positionNumberR & SUBSEPARATOR & playerNameR & SUBSEPARATOR & classNumberR & SUBSEPARATOR & pvpLevelR & SUBSEPARATOR & pvpExpR & inside

	Call Ok(str)
%>
<%
function UpdatePvpLevel(fileName)
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
		.SPName = "dbo.Rank_PvpLevel"
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
		Set RS = sph.rs
	End If
	set sph = Nothing

	If TypeName(RS)<>"Recordset" Then Call Error("no recordset")
	i = 0
objFile.WriteLine("<rank>" )
	Do Until RS.Eof
		i=i+1
		objFile.WriteLine("	<player>" )
		objFile.WriteLine("		<positionNumber>" & i & "</positionNumber>" )
		objFile.WriteLine("		<characterName>" & rs(0) & "</characterName>" )
		objFile.WriteLine("		<pvpLevel>" & rs(1) & "</pvpLevel>" )
		objFile.WriteLine("		<pvpExp>" & rs(2) & "</pvpExp>" )
		objFile.WriteLine("		<classNumber>" & rs(3) & "</classNumber>" )
		objFile.WriteLine("	</player>" )
		rs.MoveNext
	Loop
objFile.WriteLine("</rank>" )
	RS.Close
end Function
%>