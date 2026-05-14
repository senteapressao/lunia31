<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,questHash
	' response : questHash

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName,questHash
	
	If UBound(Parameters)<1 Then Call Error("not enough parameter")
	characterName = Parameters(0)

	For i=1 To Ubound(Parameters)
		questHash = Parameters(i)
		
		' call Stored procedure
		Set sphn = new SPHelper_NoTran
		with sphn
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = characterDBconnectionString
			.SPName = "dbo.CompletedActiveItemQuest"
			Call .InitCommand()
			Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
			Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
			Call .AppendParam("@questHash",adInteger,adParamInput,,questHash)
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

		If Not(rs.Eof) Then
			If retString<>"" Then
				retString = retString & SEPARATOR
			End If
			
			retString = retString & rs(0)
		End If
		rs.close
	Next

	Call Ok(retString)
%>