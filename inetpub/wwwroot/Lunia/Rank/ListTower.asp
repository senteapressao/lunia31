<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/helpers.asp"-->
<!--#include virtual="/include/common.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%  
	'Parameters = Split("0191410020wk999901",SEPARATOR)	
	Dim nRankSearch, playerName, pageNumber, playerSearch, classSearch
	If UBound(parameters)<5 Then Call Error("not enough parameter")
	nRankSearch = Parameters(0)
	playerName	= Parameters(1)
	searchParam1 = Parameters(2)
	searchParam2 = Parameters(4)
	classSearch = Parameters(5)
	playerSearch = Parameters(6)
	pageNumber = CInt(Parameters(8)) 
	if pageNumber = 0 then 'In case is the first search lunia sends in pageNumber = 0
		pageNumber = 1
	end if
	
	dim filename
	filename = CStr(getCacheFile())
	
	dim fs,f
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(filename) then
		set f=fs.GetFile(filename)
		if DateAdd("h",1,CDate(f.DateLastModified)) < Now() then
			UpdateListTower(filename)
		end if
		set f=nothing
	end if
	set fs=nothing

	Set objXMLDoc = Server.CreateObject("MSXML2.DOMDocument.6.0") 
	objXMLDoc.async = False    
	objXMLDoc.load filename
	
	dim i : i = 1
	dim positionNumberR : positionNumberR = "-"
	dim playerNameR : playerNameR = "-"
	dim classNumberR  : classNumberR = 0
	dim floorR : floorR = 0
	dim playTimeR : playTimeR = 0
	
	dim retString : retString = ""
	for each item in objXMLDoc.selectNodes("/rank/"&searchParam2&"/player") 
		'dim positionNumber : positionNumber = item.selectSingleNode("positionNumber").text 
		dim characterName : characterName 	= item.selectSingleNode("characterName").text
		dim classNumber : classNumber 		= item.selectSingleNode("classNumber").text 
		dim floor : floor 					= item.selectSingleNode("floor").text
		dim playTime : playTime 			= item.selectSingleNode("playTime").text
		dim stageGroupHash : stageGroupHash = item.selectSingleNode("stageGroupHash").text
		
		if ((classSearch = "9999" or classSearch = classNumber) and stageGroupHash=searchParam1) then
			if nRankSearch > 0 or playerSearch <> "" then
				if nRankSearch = i or InStr(LCase(characterName),playerSearch) > 0 then
					retString = retString &_
						SEPARATOR 	 & i &_
						SUBSEPARATOR & characterName &_
						SUBSEPARATOR & classNumber &_
						SUBSEPARATOR & floor &_
						SUBSEPARATOR & playTime
					j = j + 1 ' items in the page
					if nRankSearch = i then Exit For end if
				end if
			elseif roundUp(i/ranksPerPage) = pageNumber then
				retString = retString &_
					SEPARATOR 	 & i &_
					SUBSEPARATOR & characterName &_
					SUBSEPARATOR & classNumber &_
					SUBSEPARATOR & floor &_
					SUBSEPARATOR & playTime
				j = j + 1 ' items in the page
			end if
			i = i + 1
		end if
	Next
	TotalPageNumber = roundUp(i/20)
	if TotalPageNumber = 0 then TotalPageNumber = 1 end if
	
	If Cint(pageNumber) > Cint(TotalPageNumber) then Call Error(Cint(pageNumber) & ">" & Cint(TotalPageNumber))
	myData = positionNumberR & SUBSEPARATOR & playerNameR & SUBSEPARATOR & classNumberR & SUBSEPARATOR & FloorR & SUBSEPARATOR & playTimeR
	Call Ok(TotalPageNumber & SEPARATOR & pageNumber & SEPARATOR & myData & retString)
%>

<%
function UpdateListTower(fileName)
	Server.ScriptTimeout = 200
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
		.SPName = "dbo.Rank_ListTower"
		Call .InitCommand()
		blsResult = .ExecRecordset4()
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
		CacheMajor RS, "dd", objFile
		Set RS = sph.rs2 : If TypeName(RS)<>"Recordset" Then Call Error("no recordset2")
		CacheMajor RS, "wk", objFile
		Set RS = sph.rs3 : If TypeName(RS)<>"Recordset" Then Call Error("no recordset3")
		CacheMajor RS, "mm", objFile
	objFile.WriteLine("</rank>")
	End If
	
	set sph = Nothing

	RS.Close
end function

	function CacheMajor(RS, majorType, objFile)
objFile.WriteLine("<"&majorType&">")
	Do Until RS.Eof
		objFile.WriteLine("	<player>" )
		objFile.WriteLine("		<characterName>" & rs(0) & "</characterName>" )
		objFile.WriteLine("		<classNumber>" & rs(1) & "</classNumber>" )
		objFile.WriteLine("		<floor>" & rs(2) & "</floor>" )
		objFile.WriteLine("		<playTime>" & rs(3) & "</playTime>" )
		objFile.WriteLine("		<stageGroupHash>" & rs(4) & "</stageGroupHash>" )
		objFile.WriteLine("	</player>" )
		rs.MoveNext
	Loop
objFile.WriteLine("</"&majorType&">")
	end function
%>