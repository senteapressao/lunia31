<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init() 
	' request	: characterName,serverName
	' response	: characterSerial,completedAchivos,statisticList
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""
	
	' preparing request parameters
	Dim characterName,serverName
	If UBound(parameters)<0 Then Call Error("not enough parameter")	
	characterName	= Parameters(0)
	serverName		= REQUESTERID
	
	Dim statisticsValue
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.achievement_InitPlayer"
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

	If rs.Eof Then Call Error(1)

    Dim characterSerial,classType

	characterSerial = rs(0)
	classType = rs(1)
	'Inserting PlayerInfo into achievements connection
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_JoinAchievement_v2"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@serverName",adVarWChar,adParamInput,50,serverName)
		Call .AppendParam("@characterSerial",adBigInt,adParamInput,,characterSerial)
		blsResult = .ExecNoRecords()
	End with

	If blsResult = False Then
		Dim temp_sph_frk_n4ErrorCode,temp_sph_frk_strErrorText
		temp_sph_frk_n4ErrorCode	= sph.frk_n4ErrorCode
		temp_sph_frk_strErrorText	= sph.frk_strErrorText
		Call Error(temp_sph_frk_n4ErrorCode)
	End If
	set sphn = Nothing
	'Done Inserting PlayerInfo into achievements connection
	
	Response.Write characterSerial & SEPARATOR & classType

	'completedAchivos
    i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0) & SUBSEPARATOR & rs(1) & SUBSEPARATOR & rs(2) & SUBSEPARATOR & FormatDt(rs(3),"SQL_TM")
		rs.MoveNext : i=i+1
	Loop
	'workingAchivos
    i = 0
	Response.Write SEPARATOR
	set rs = rs.NextRecordset
	Do Until rs.Eof
		If i>0 Then Response.Write SUBSEPARATOR
		Response.Write rs(0) & SUBSEPARATOR & rs(1)
		rs.MoveNext : i=i+1
	Loop

%>