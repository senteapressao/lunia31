<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request : tax
	' response : guildId,tax,tax_o
	
	Dim i,j,k
	Dim sph,rs,retString,blsResult : blsResult = False

	' preparing request parameters
	Dim tax
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	tax = Parameters(0)
	
	Dim guildId,guildName,tax_o
	tax_o = tax
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_PutTax"
		Call .InitCommand()
		Call .AppendParam("@tax",adInteger,adParamInputOutput,,tax)
		Call .AppendParam("@guildId",adInteger,adParamOutput,,null)
		Call .AppendParam("@guildName",adVarWChar,adParamOutput,14,null)
		blsResult = .ExecRecordset()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		tax = sph.GetParamValue("@tax")
		guildId = sph.GetParamValue("@guildId")
		guildName = sph.GetParamValue("@guildName")
	End If
	Set sph = Nothing
	
	retString = retString & guildId & SEPARATOR & guildName & SEPARATOR & tax & SEPARATOR & tax_o

	Call Ok(retString)
%>