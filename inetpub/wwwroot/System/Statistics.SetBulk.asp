<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request	: (POST)	
	'			characterSerial
	'				statisticsKey,statisticsValue
	' response	:
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	
	Dim rowData
	rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)
	
	' preparing request parameters
	Dim characterSerial
	If UBound(parameters)<1 Then Call Error("not enough parametergg")	
	characterSerial	= Parameters(0)
	
	dim values,arr
	' build query
	values = ""
	arr = Split(Parameters(1),SUBSEPARATOR)
	
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 2)<>0 Then
			Call Error("invalid number of statistics parameters")
		End If
		
		For i=0 To UBound(arr) Step 2
			If i <> 0 Then
				values = values & "," 
			End If
			values = values & "('" & arr(i) & "','" & arr(i+1) & "')"
		Next
		Dim sql
		sql = "UPDATE statisticsList " &_
			"SET " &_
			"statisticsValue  = t1.statistic_value " &_
			"FROM (" &_
			"VALUES" &_
			values &_
			") AS t1 (statistic_key,statistic_value) " &_
			"WHERE characterSerial = " & characterSerial & " and t1.statistic_key = statisticsKey"
		Dim conn
		Set conn = Server.CreateObject("ADODB.Connection")
		conn.Open(characterDBconnectionString)
		'conn.Execute(sql)
		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.Open sql,conn,adOpenStatic,adLockReadOnly,adCmdText
		set rs = rs.NextRecordset
		conn.Close()
		Call Ok(0)
	End If
	Call Ok(0)
%>