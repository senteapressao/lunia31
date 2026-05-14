<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : character name, InviterName, stagegrouphash , accesslevel
	' response : stage server address, port, stage group hash, access level of the stage group
	
	Dim i,j,k
	Dim sph,rs,blsResult : blsResult = False

	' preparing request parameters
	Dim characterName,	inviterName, stageGroupHash, accessLevel ,teamNumber
	Dim	address,port,reservationSerial
	response.write UBound(parameters)
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	characterName = Parameters(0)
	inviterName = Parameters(1)
	
	' TODO: parameter validation

	' call Stored procedure	
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = stageDBconnectionString
		.SPName = "dbo.public_Spectate"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@inviterName",adVarWChar,adParamInput,50,inviterName)
		Call .AppendParam("@address",adVarChar,adParamOutput,50,null)
		Call .AppendParam("@port",adInteger,adParamOutput,,null)
		Call .AppendParam("@reservationSerial",adBigInt,adParamOutput,,null)
		blsResult = .ExecNoRecords()
	End with

	If blsResult=False Then
		if sph.frk_n4ErrorCode>0 Then
			Call Error(sph.frk_n4ErrorCode)
		Else
			Call Error(sph.frk_strErrorText)
		End If
	Else
		address = sph.GetParamValue("@address")
		port = sph.GetParamValue("@port")
		reservationSerial = sph.GetParamValue("@reservationSerial")
	End If
	set sph = Nothing

	Call Ok(address & SEPARATOR & port & SEPARATOR & reservationSerial)
%>
