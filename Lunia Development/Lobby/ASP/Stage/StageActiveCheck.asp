<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	'Logger()
	Call Init()
	' request	: room number
	' response	: -

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	Dim serverName
	serverName = REQUESTERID
	dim values
	' build query
	values = ""
	
	If Ubound(Parameters)>=0 Then
		
		For i=0 To UBound(Parameters) Step 1
			If i <> 0 Then
				values = values & "," 
			End If
			values = values & "(" & Parameters(i) & ")"
		Next
		Dim sql
		sql = "INSERT INTO [dbo].[InActiveStages](ServerName,roomNumber) " &_
   			"select '" & REQUESTERID & "', t1.roomNumber " &_
   				"FROM ( " &_
   					"VALUES " &_
   						values &_
				") AS t1 (roomNumber) " &_
   			"WHERE NOT EXISTS( " &_
   				"SELECT * FROM [dbo].[InActiveStages] as t2 " &_
   				"WHERE " &_
   					"t2.ServerName = '" & REQUESTERID & "' and " &_
   					"t2.roomNumber=t1.roomNumber " &_
			")"
		Dim conn
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open(stageDBconnectionString)
		'conn.Execute(sql)
		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.Open sql,conn,adOpenStatic,adLockReadOnly,adCmdText
		set rs = rs.NextRecordset
		conn.Close()
		Call Ok(0)
	End If
	Call Ok(0)
%>