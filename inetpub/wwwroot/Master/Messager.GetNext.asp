<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	' request	: (empty)
	' response	: command,message, toSquare,toStage
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim rs2
	Dim retString : retString = ""
	' preparing request parameters
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = gateDBconnectionString
		.SPName = "dbo.public_NextMessage"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		blsResult = .ExecRecordset()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")		
		if ret<>0 Then
			Call Error(ret)
		End If
		Set rs = sphn.rs
		Set sphn = Nothing

		i=0
		Do Until rs.Eof
			retString = retString & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2)
			rs.MoveNext:i=i+1
		Loop
		rs.close
	Else
		retString = 0
	End If
	

	Call Ok(retString)
%>