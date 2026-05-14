<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/helpers.asp"-->
<!--#include virtual="/include/common.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : virtualIdCode
	' response : characterName,classNumber,stageLevel,stageExp,pvpLevel,pvpExp
	
	Dim i,j,k
	Dim sph,rs,rs2,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim virtualIdCode
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	virtualIdCode = Parameters(0)
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.ListCharacter"
		Call .InitCommand()
		Call .AppendParam("@virtualIdCode","@characterName","@classNumber","@stageLevel","@pvpLevel",,virtualIdCode)
		blsResult = .ExecRecordset()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		Set rs = sph.rs
	End If
	Set sph = Nothing
	
	i = 0
	Do Until rs.Eof
		If i>0 Then retString = retString & SEPARATOR
		retString = retString & rs(0) & SUBSEPARATOR & rs(1) &_
				SUBSEPARATOR & rs(2) & SUBSEPARATOR & rs(3) &_
				SUBSEPARATOR & rs(4) & SUBSEPARATOR & rs(5) &_
				SUBSEPARATOR & rs(6)
				
		rs.MoveNext : i=i+1
	Loop
	rs.close
	
	Call Ok(retString)
%>