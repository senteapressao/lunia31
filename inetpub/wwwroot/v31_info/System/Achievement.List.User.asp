<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request	: achievementHash
	' response	: achievementHash,{characterName}
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim rs2
	Dim retString : retString = ""
	'459018721
	' preparing request parameters
	Dim achievementHash, isFirst, achievementCount
	If UBound(parameters)<1 Then Call Error("not enough parameter")	
	achievementHash	= Parameters(0)
	isFirst =  Parameters(1)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.achievementListUser"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@achievementHash",adBigInt,adParamInput,,achievementHash)
		Call .AppendParam("@isFirst",adUnsignedTinyInt,adParamInput,,isFirst)
		Call .AppendParam("@achievementCount",adInteger,adParamOutput,,null)
		blsResult = .ExecRecordset()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		
		achievementCount = sphn.GetParamValue("@achievementCount")
		
		if ret<>0 Then
			Call Error(ret)
		End If
		Set rs = sphn.rs
	Else
		Call Error(1)
	End If
	set sphn = Nothing

	retString = achievementHash & SEPARATOR
	

	i=0
	Do Until rs.Eof
		If i>0 Then
			retString = retString & SUBSEPARATOR & ", "
		End If
		
		retString = retString & rs(0)
	

		rs.MoveNext:i=i+1
	Loop
	rs.close
	
	retString = retString & " [br][br]" & SEPARATOR & achievementCount

	Call Ok(retString)
%>