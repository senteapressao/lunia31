<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	' response : mailId,sender,title,status,sendDate
	
	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterName
	characterName = Parameters(0)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.ReturnMail"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecNoRecords()
	End with
	set sphn = Nothing
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.DeleteReturnMail"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecNoRecords()
	End with
	set sphn = Nothing
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.ListMail"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset()
	End with
	
	If blsResult Then
		Set rs = sphn.rs
	End If
	set sphn = Nothing
	
	i=0
	Do Until rs.Eof
		If i>0 Then
			retString = retString & SEPARATOR
		End If
			
		retString = retString & rs(0) & SUBSEPARATOR & rs(1) &_
				SUBSEPARATOR & rs(2)
		If rs(3) Then
			retstring = retstring & SUBSEPARATOR & 1
		Else
			retstring = retstring & SUBSEPARATOR & 0
		End If
		retString = retString & SUBSEPARATOR & rs(4) & SUBSEPARATOR & FormatDt(rs(5),"SQL_TM")

		rs.MoveNext:i=i+1
	Loop
	rs.close
	
	Call OK(retString)
%>