<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : accountName, password, newPassword
	' response : isSuccess? [0:1:2] || 2 = fail, but due to new password being the same as current, and not because of [current] password not matching.
	
	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim rs2,rs3
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterName, password, newPassword
	characterName=Parameters(0)
	password=Parameters(1)
	newPassword=Parameters(2)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.SecondPassword_Modify"
		Call .InitCommand()
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@password",adVarWChar,adParamInput,50,password)
		Call .AppendParam("@newPassword",adVarWChar,adParamInput,50,newPassword)
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
			
		retString = retString & rs(0)
		rs.MoveNext:i=i+1
	Loop
	rs.close
	
	Call OK(retString)
%>
