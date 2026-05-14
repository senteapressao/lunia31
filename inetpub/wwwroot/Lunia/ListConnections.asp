<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Ok("0" & SUBSEPARATOR & 10 & SUBSEPARATOR & 3000 & SEPARATOR)
	
	' Write to file if greater
	dim filename,fs,f,a,currentcount,maxcount
	Dim conn,rs
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open(stageDBconnectionString)
	rs = conn.Execute("select count('x') from dbo.connections (nolock)")
	currentcount = CInt(rs(0))
	conn.Close()
	
	filename = replace(Server.MapPath("./"),"\","/") & "/" & "ListConnection.cache"
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	If fs.FileExists(filename) then
		set a = fs.GetFile(filename)
		maxcount = Cint(a.OpenAsTextStream(1,-2).ReadLine)
		If currentcount > maxcount Then
			UpdateCount fs,currentcount
			maxcount = currentcount
		end If
	Else
		'Create file
		set a=fs.OpenTextFile(filename, 2, true, -2) 
		a.WriteLine(currentcount)
		a.Write(now)
		a.Close()
		maxcount = currentcount
	End If
	
	' currentcount - X
	' maxcount 	   - 13
	Call Ok("0" & SUBSEPARATOR & CInt((currentcount*13/maxcount)) & SUBSEPARATOR & 3000 & SEPARATOR)
	
	Function UpdateCount(filesystem,count)
		dim a_file
		Set a_file = filesystem.OpenTextFile(2, -2) 
		a_file.WriteLine(count)
		a_file.Write(now)
		a_file.Close
	End Function
%>