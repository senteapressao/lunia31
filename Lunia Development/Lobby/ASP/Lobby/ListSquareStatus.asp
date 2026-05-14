<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	'If UBound(parameters)<3 Then Call Error("not enough parameter")

	' request  : 
	' response : {squareName,isBlocked,connectionCount,capacity,stageGroupHash,accessLevel,orderNumber}

	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_ListSquareStatus"
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
		retString = retString & rs(0) & SUBSEPARATOR & 0 &_
				SUBSEPARATOR & rs(1) & SUBSEPARATOR & rs(2) & SUBSEPARATOR & rs(3) &_
				SUBSEPARATOR & rs(4) & SUBSEPARATOR & rs(5)
		rs.MoveNext : i=i+1
	Loop
	RS.Close

	Call Ok(retString)
%>