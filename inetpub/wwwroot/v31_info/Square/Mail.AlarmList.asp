<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : (POST)
	'			characterName
	' response : characterName
	
	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	Dim rowData
	rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)
	
	If Ubound(Parameters)<0 Then Call OK(retString)
	
	Dim conn,sql
	set conn = Server.CreateObject("ADODB.Connection")
	conn.Open(characterDBconnectionString)
	
	sql = "select distinct characterName from dbo.Mails (nolock)"&_
		" where isRead=0 and characterName in ("

	For i=0 To Ubound(Parameters)
		If i>0 Then sql = sql &","
		sql = sql & SqlQuot(Parameters(i))
	Next
	
	sql = sql &")"
	sql = sql &" and ( state > 0 OR isSelf = 1 OR datediff(ss,senddate,getdate()) > 3600 )"
	
	set rs = conn.Execute(sql)
	
	i=0
	Do Until rs.Eof
		If i>0 Then
			retString = retString & SEPARATOR
		End If
			
		retString = retString & rs(0)

		rs.MoveNext:i=i+1
	Loop
	rs.close
	conn.Close
	
	Call OK(retString)
%>