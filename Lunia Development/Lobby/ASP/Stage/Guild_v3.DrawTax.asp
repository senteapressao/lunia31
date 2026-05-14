<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request : guildId,guildMemberId,tax
	' response : guildId,tax,tax_o
	
	Dim i,j,k
	Dim sph,rs,retString,blsResult : blsResult = False

	' preparing request parameters
	Dim guildId,guildMemberId,tax
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	guildId = Parameters(0)
	guildMemberId = Parameters(1)
	tax = Parameters(2)
	
	Dim taxPayDate,tax_o
	tax_o = tax
	
	' call Stored procedure
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.public_DrawTax"
		Call .InitCommand()
		Call .AppendParam("@guildId",adInteger,adParamInput,,guildId)
		Call .AppendParam("@guildMemberId",adInteger,adParamInput,,guildMemberId)
		Call .AppendParam("@tax",adInteger,adParamInputOutput,,tax)
		Call .AppendParam("@taxPayDate",adDBTimeStamp,adParamOutput,,null)
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
		taxPayDate = sph.GetParamValue("@taxPayDate")
	End If
	Set sph = Nothing
	
	retString = retString & guildId & SEPARATOR & tax & SEPARATOR & FormatDt(taxPayDate,"SQL_TM") & SEPARATOR & tax_o

	Call Ok(retString)
%>