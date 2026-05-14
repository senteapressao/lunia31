<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,pageIndex
	' response : pageIndex
	'			,{itemSerial,itemHash,stackedCount,instance,position}

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	Dim characterName,pageIndex
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	charactername = Parameters(0)
	pageIndex = Parameters(1)
	
	Dim sql
	sql = "select itemSerial,itemHash,stackedCount,instance,position from dbo.CashItemStorage (nolock)"&_
		" where characterName=? and pageIndex=?"

	with Command
		.ActiveConnection = characterDBconnectionString
		.CommandText = sql
		.CommandType = adCmdText
		
		.Parameters.Append .CreateParameter("@characterName",adVarWChar,adParamInput,50,characterName)
		.Parameters.Append .CreateParameter("@pageIndex",adUnsignedTinyInt,adParamInput,,pageIndex)		
	End with
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.CursorLocation = adUseClient
	rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdText
	
	If rs.State = adStateClosed Then
		Call Error(1)
	End If
	
	retString = retString & pageIndex
	
	i=0
	Do Until rs.Eof
		retString = retString & SEPARATOR & rs(0) & SUBSEPARATOR & rs(1) & SUBSEPARATOR & rs(2) & SUBSEPARATOR & rs(3) & SUBSEPARATOR & rs(4)

		rs.MoveNext:i=i+1
	Loop
	rs.close
	
	Call Ok(retString)
%>
