<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,petSerial,petName,rareProbability,petHash,petEnchantSerial
	' response : petSerial,petName,petLevel,petExp,fullValue,isRare,rareProbaliity,petHash
		
	Dim sphn,blsResult : blsResult = False
	Dim retString : retString = ""	

	Dim characterName,petSerial,petName,rareProbability,petEnchantSerial
	Dim petLevel,petExp,fullValue,isRare,petHash
	characterName	= Parameters(0)
	petSerial		= Parameters(1)
	petName			= Parameters(2)
	rareProbability	= CDbl(Parameters(3))
	petHash			= Parameters(4)
	petEnchantSerial= Parameters(5)
	
	petLevel = 1
	petExp = 0
	fullValue = 0
	isRare = 0
	
	' call Stored procedure	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.CreatePet"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		Call .AppendParam("@petSerial",adBigInt,adParamInput,,petSerial)
		Call .AppendParam("@petName",adVarWChar,adParamInput,50,petName)		
		Call .AppendParam("@rareProbability",adSingle,adParamInput,,rareProbability)
		Call .AppendParam("@petHash",adInteger,adParamInput,,petHash)
		Call .AppendParam("@petEnchantSerial",adInteger,adParamInput,,petEnchantSerial)
						
		blsResult = .ExecNoRecords()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
	End If
	set sphn = Nothing
	
	Call Ok(petSerial & SEPARATOR & petName & SEPARATOR & petLevel & SEPARATOR & petExp & SEPARATOR & fullValue & SEPARATOR & Int(isRare) & SEPARATOR & rareProbability & SEPARATOR & petHash & SEPARATOR & petEnchantSerial)
%>