<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	' response : Friend list (separated : characterName, class, level, isFriend )
	
	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim rs2,rs3
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterName
	characterName=Parameters(0)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.FriendGetList"
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
				SUBSEPARATOR & rs(2) & SUBSEPARATOR & rs(3)

		rs.MoveNext:i=i+1
	Loop
	rs.close
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.FriendGetReferencedList"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset()
	End with
	
	If blsResult Then
		Set rs = sphn.rs
	End If
	set sphn = Nothing
	
	Do Until rs.Eof
		If retString<>"" Then
			retString = retString & SEPARATOR
		End If
		retString = retString & rs(0)
		
		rs.MoveNext
	Loop
	rs.close
	
	Call OK(retString)
%>
