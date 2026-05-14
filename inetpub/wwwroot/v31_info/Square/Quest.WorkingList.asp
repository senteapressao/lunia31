<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	' response : questHash,currentState,param1,param2,param3,expiredDate

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName
	
	If UBound(Parameters)<0 Then Call Error("not enough parameter")
	characterName = Parameters(0)

	' call Stored procedure
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.ListWorkingQuests"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		Set rs = sphn.rs
	Else
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
	End If
	set sphn = Nothing

	i=0
	Do Until rs.Eof
		If i>0 Then
			retString = retString & SEPARATOR
		End If
			
		retString = retString & rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) &_
				SEPARATOR & rs(3) & SEPARATOR & rs(4) & SEPARATOR & FormatDt(rs(5),"SQL_TM")

		rs.MoveNext:i=i+1
	Loop
	rs.close

	Call Ok(retString)
%>