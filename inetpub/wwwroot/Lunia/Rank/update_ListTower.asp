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
		CacheMajor RS, "dd"
		Set RS = sph.rs2 : If TypeName(RS)<>"Recordset" Then Call Error("no recordset2")
		CacheMajor RS, "wk"
		Set RS = sph.rs3 : If TypeName(RS)<>"Recordset" Then Call Error("no recordset3")
		CacheMajor RS, "mm"
	objFile.WriteLine("</rank>")
	End If
	
	set sph = Nothing

	RS.Close
	Call Ok("Updated")
%>
<%
	function CacheMajor(RS, majorType)
objFile.WriteLine("<"&majorType&">")
	i = 0
	Do Until RS.Eof
		i=i+1
		objFile.WriteLine("	<player>" )
		objFile.WriteLine("		<positionNumber>" & i & "</positionNumber>" )
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