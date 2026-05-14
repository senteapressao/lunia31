<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : -
	' response : square list (square name, isBlocked, population, capacity)

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_ListPvpChannelStatus"
		Call .InitCommand()
		blsResult = .ExecRecordset()
	End with

	If blsResult=False Then		
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		Set RS = sph.rs
	End If
	set sph = Nothing

	If TypeName(RS)<>"Recordset" Then Call Error("no recordset")
	i = 0
	Do Until RS.Eof
		If i>0 Then
			retString = retString & SEPARATOR
		End If
		retString = retString & rs(0) & SUBSEPARATOR & rs(1) &_
				SUBSEPARATOR & rs(2) & SUBSEPARATOR & rs(3) & SUBSEPARATOR & rs(4)
		rs.MoveNext : i=i+1
	Loop
	RS.Close

	Call Ok(retString)
%>