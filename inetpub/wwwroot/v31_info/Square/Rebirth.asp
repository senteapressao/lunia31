<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	
	' request  : characterName,rebirthCount,storedLevel,storedSkillPoint
	' response : rebirthCount,storedLevel,storedSkillPoint,lastRebirthDate
	'			,licenses(subseparated: stageGroupHash,accessLevel)
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	Dim characterName,rebirthCount,storedLevel,storedSkillPoint
	
	characterName = Parameters(0)
	rebirthCount = Parameters(1)
	storedLevel = Parameters(2)
	storedSkillPoint = Parameters(3)
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.UpdateRebirth"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@rebirthCount",adSmallInt,adParamInput,,rebirthCount)
		Call .AppendParam("@storedLevel",adSmallInt,adParamInput,,storedLevel)
		Call .AppendParam("@storedSkillPoint",adSmallInt,adParamInput,,storedSkillPoint)
		blsResult = .ExecNoRecords()
	End with
	set sphn = Nothing
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.InfoRebirth"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset()
	End with
	
	If blsResult Then
		Set rs = sphn.rs
	Else
		Call Error("no recordset")
	End If
	set sphn = Nothing
	
	' rebirthCount,storedLevel,storedSkillPoint,lastRebirthDate
	retString = rs(0) & SEPARATOR & rs(1) & SEPARATOR & rs(2) & SEPARATOR & FormatDt(rs(3),"SQL_TM")
	
	' licenses(subseparated: stageGroupHash,accessLevel)
	i=0
	retString = retString & SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then
			retString = retString & SUBSEPARATOR
		End If
		retString = retString & rs(0) & SUBSEPARATOR & rs(1)
		rs.MoveNext : i=i+1
	Loop
	rs.Close

	Call Ok(retString)
%>