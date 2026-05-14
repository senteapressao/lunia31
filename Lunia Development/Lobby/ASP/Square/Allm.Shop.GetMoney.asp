<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init() 
	' request	: characterSerial
	' response	: {achievementHash(complete,logDate,isShow,isSend)}
	'			,{achievementHash(woring)}
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim rs2
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterName,accountName
	If UBound(parameters)<0 Then Call Error("not enough parameter")	
	characterName	= Parameters(0)
	accountName = GetAccountName(characterName)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = cashDBconnectionString
		.SPName = "dbo.Acc_GetCash"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@accountName",adVarWchar,adParamInput,50,accountName)
		blsResult = .ExecRecordset()
	End with
	
	If blsResult Then
		Set rs = sphn.rs
	Else
		Call Error(1)
	End If
	set sphn = Nothing
	
	dim nCash,nPoints
	i=0
	Do Until rs.Eof
		If i = 0 Then
			nCash = rs(0)
			nPoints = rs(1)
		End If

		rs.MoveNext:i=i+1
	Loop
	rs.close
	Call Ok(characterName & SEPARATOR & accountName & SEPARATOR & nCash & SEPARATOR & nPoints)
%>