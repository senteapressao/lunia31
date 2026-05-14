<!--#include virtual="/include/connect.asp"-->
<!--#include virtual="/include/common.asp"-->
<!--#include virtual="/include/helpers.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	' request  : 
	' response : characterName,accountName,virtualIdCode,classNumber,stageLevel,stageExp,pvpLevel,pvpExp,instantStateFlags,warLevel,warExp
	
	If UBound(parameters)<1 Then Call Error("not enough parameter")
	classNumber = Parameters(0)
	characterName = Parameters(1)
	
	Dim i,j,k
	Dim sph,blsResult : blsResult = False
	Dim retString : retString = ""

	' call Stored procedure
	Set sph = new SPHelper_NoTran
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = ConnStrCharacter
		.SPName = "dbo.CharacterInfoForPvp"
		Call .InitCommand()
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
		blsResult = .ExecRecordset()
	End with

	If blsResult Then		
		Set rs = sph.rs
	End If
	set sph = Nothing

	
	If TypeName(rs)<>"Recordset" Then Call Error("no recordset")
	If Not(rs.Eof) Then
		retString = retString &_
		 	rs(0) & SEPARATOR &_					
		 	rs(1) & SEPARATOR &_
			rs(2) & SEPARATOR &_
			(cint(rs(1)) - cint(rs(2))) & SEPARATOR &_
			
			rs(3) & SEPARATOR &_
			rs(4) & SEPARATOR &_
			rs(6) & SEPARATOR &_
			(cint(rs(5)) - cint(rs(6))) & SEPARATOR &_
			rs(8) & SEPARATOR &_
			rs(9) & SEPARATOR &_
			0 & SEPARATOR &_
			0
	Else
		retString = retString &_
		 	1500 & SEPARATOR &_ 
		 	0 & SEPARATOR &_
			0 & SEPARATOR &_
			0 & SEPARATOR &_
			
			0 & SEPARATOR &_
			0 & SEPARATOR &_
			0 & SEPARATOR &_
			0 & SEPARATOR &_
			
			0 & SEPARATOR &_
			0 & SEPARATOR &_
			0 & SEPARATOR &_
			0
	End If

	' data.LadderPoint = StringUtil::To< uint32 >( dataStrings.at( 0 ) );
	' data.LadderMatchCount = static_cast< uint16 >( StringUtil::To< uint32 >( dataStrings.at( 1 ) ) );
	' data.LadderWinCount = StringUtil::To< uint32 >( dataStrings.at( 3 ) );
	' data.LadderLoseCount = StringUtil::To< uint32 >( dataStrings.at( 2 ) ) - data.LadderWinCount;

	' data.PvpLevel = StringUtil::To< uint32 >( dataStrings.at( 4 ) );
	' data.PvpExp = StringUtil::To< uint32 >( dataStrings.at( 5 ) );
	' data.PvpWinCount = StringUtil::To< uint32 >( dataStrings.at( 6 ) );
	' data.PvpLoseCount = StringUtil::To< uint32 >( dataStrings.at( 7 ) );
	
	' data.WarLevel = StringUtil::To< uint32 >( dataStrings.at( 8 ) );
	' data.WarExp = StringUtil::To< uint32 >( dataStrings.at( 9 ) );
	' data.WarWinCount = StringUtil::To< uint32 >( dataStrings.at( 10 ) );
	' data.WarLoseCount = StringUtil::To< uint32 >( dataStrings.at( 11 ) );
	
	Call Ok(retString)
%>