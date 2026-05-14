<!--#include file="../include/connect.asp"-->
<!--#include file="../include/common.asp"-->
<!--#include file="../include/helpers.asp"-->
<%
	' fileName for the cache file.
	dim fs,destinationPath
	destinationPath = getCacheFile()
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
		CacheMajor RS, "achievement"
	objFile.WriteLine("</rank>")
	End If
	
	set sph = Nothing

	RS.Close
	Call Ok("Updated")
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
%>