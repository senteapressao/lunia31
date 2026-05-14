<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,bagNumber
	' response : characterName,bagNumber
	'			,{positionNumber,itemHash,stackedCount,itemSerial,instance,itemExpire}

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName,bagNumber
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	characterName	= Parameters(0)
	bagNumber		= Parameters(1)

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.BankItems_List"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@bagNumber",adUnsignedTinyInt,adParamInput,,bagNumber)
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
		
		retString = retString & rs(1) & SUBSEPARATOR & rs(2) & SUBSEPARATOR & rs(3) &_
				SUBSEPARATOR & rs(4) & SUBSEPARATOR & rs(5) & SUBSEPARATOR & FormatDt(rs(6),"SQL_TM")

		rs.MoveNext:i=i+1
	Loop
	rs.close

	retString = characterName & SEPARATOR & bagNumber & SEPARATOR & retString

	Call OK(retString)
%>