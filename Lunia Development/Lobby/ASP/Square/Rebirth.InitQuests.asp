<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	
	' request  : characterName,isQuestReset
	' response : characterName
	
	Dim i,j,k
	Dim sph,sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	Dim characterName,isQuestReset
	
	characterName = Parameters(0)
	isQuestReset = 1

	If Ubound(Parameters) > 0 Then
		isQuestReset = Parameters(1)
	end If

	if isQuestReset = 1 Then
		Set sphn = new SPHelper_NoTran
		with sphn
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = characterDBconnectionString
			.SPName = "dbo.Rebirth_InitQuests"
			Call .InitCommand()
			Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
			blsResult = .ExecNoRecords()
		End with
		set sphn = Nothing
	End If
	Call Ok(characterName)
%>