<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()	
	' request  : characterName
	' response : characterName,rank
	' RankPvpWeekly Reward Buff
	
	Dim sphn,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName, accountName
	'If UBound(parameters)<0 Then Call Error("not enough parameter")
	characterName	= Parameters(0)
	'accountName = GetAccountName(characterName)
	
	
	'Dim rank
	
	'call Stored procedure	
	Dim ret,rs
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.GetRewardRank"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset()
	End with
	
	If blsResult Then
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		Set rs = sphn.rs
		
		ret = ""
		Do Until rs.Eof
			ret = ret & SEPARATOR & rs(0)
			rs.MoveNext
		Loop
		rs.Close
		
	Else
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
	End If
	
	set sphn = Nothing
	
	Call Ok(characterName & ret)
	
%>