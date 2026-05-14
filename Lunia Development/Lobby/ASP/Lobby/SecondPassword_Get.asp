<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : accountName
	' response : isUsingSecondPassword, failedAttemptCount, lockExpireDate, isLocked
	
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
		.SPName = "dbo.SecondPassword_Get"
		Call .InitCommand()
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,characterName)
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
			
		retString = retString & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & rs(3) 
		rs.MoveNext:i=i+1
	Loop
	rs.close
	
	Call OK(retString)
%>
