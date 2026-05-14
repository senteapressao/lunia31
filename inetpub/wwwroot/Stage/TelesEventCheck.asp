<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : stageHash
	' response : state,{name,probability}
	
	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim stageHash,dropToAll
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	stageHash = Parameters(0)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.SPName = "dbo.TelesEventCheck"
		.ConnStr = characterDBconnectionString
		Call .InitCommand()
		Call .AppendParam("@stageHash",adInteger,adParamInput,,stageHash)
		Call .AppendParam("@dropToAll",adInteger,adParamOutput,,null)
		blsResult = .ExecRecordset()
	End with
	
	If blsResult Then
		Set rs = sphn.rs
		dropToAll = sphn.GetParamValue("@dropToAll")
	End If
	set sphn = Nothing
	
	i=0
	Do Until rs.Eof
		If i>0 Then
			retString = retString & SEPARATOR
		End If
		retString = retString & rs(0) & SUBSEPARATOR & rs(1)
		rs.MoveNext:i=i+1
	Loop
	rs.close
	
	Call OK(dropToAll & SEPARATOR & retString)
%>