<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : guildId
	' response : itemHash,expireDate
	
	Dim i,j,k
	Dim sph,rs,rs2,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim guildId
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	guildId = Parameters(0)
	
	Dim guildLevel
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_ListGuildShopProduct"
		Call .InitCommand()
		Call .AppendParam("@guildId",adInteger,adParamInput,,guildId)
		Call .AppendParam("@guildLevel",adUnsignedTinyint,adParamOutput,,null)
		blsResult = .ExecRecordset()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		guildLevel = sph.GetParamValue("@guildLevel")
		Set rs = sph.rs
	End If
	Set sph = Nothing
	
	retString = retString & guildLevel
	
	i = 0
	Do Until rs.Eof
		If i=0 Then
			retString = retString & SEPARATOR
		Else
			retString = retString & SUBSEPARATOR
		End If
		retString = retString & rs(0) & SUBSEPARATOR & FormatDt(rs(1),"SQL_TM")
				
		rs.MoveNext : i=i+1
	Loop
	rs.close
	
	Call Ok(retString)
%>