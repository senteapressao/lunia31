<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	' response : characterName
	
	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterName
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	
	Dim sql
	sql = "select distinct characterName from dbo.Mails (nolock)"&_
		" where isRead=0 and characterName=?"
	sql = sql &" and ( state > 0 OR isSelf = 1 OR datediff(ss,senddate,getdate()) > 3600 )"

	with Command
		.ActiveConnection = characterDBconnectionString
		.CommandText = sql
		.CommandType = adCmdText
		
		.Parameters.Append .CreateParameter("@characterName",adVarWChar,adParamInput,50,characterName)
	End with
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.CursorLocation = adUseClient
	rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdText
	
	If rs.State = adStateClosed Then
		Call Error(1)
	End If

	i=0
	Do Until rs.Eof
		If i>0 Then
			retString = retString & SEPARATOR
		End If
			
		retString = retString & rs(0)

		rs.MoveNext:i=i+1
	Loop
	rs.close
	
	Call Ok(retString)
%>