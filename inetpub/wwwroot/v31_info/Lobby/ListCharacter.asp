<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : accountName
	' response : slotCount,accountName
	'		,{classNumber}
	'		,{characterName,CharacterSerial,virtualIdCode,classNumber,stageLevel,stageExp,pvpLevel,pvpExp,warLevel,warExp,lastLoggedDate,instantStateFlags,rebirthCount,storedLevel}
	'		,{characterName,stageGroupHash,accessLevel}
	'		,{characterName,bagNumber,positionNumber,itemHash,instance,itemExpire}

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim rs2,rs3,rs4
	Dim retString : retString = ""

	' preparing request parameters
	Dim accountName
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	accountName	= Parameters(0)
	
	Dim slotCount
	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.ListCharacters"
		Call .InitCommand()
		Call .AppendParam("@accountName",adVarWChar,adParamInput,50,accountName)
		Call .AppendParam("@slotCount",adInteger,adParamInputOutput,,null)
		blsResult = .ExecRecordset4()
	End with

	If blsResult Then
		slotCount = sphn.GetParamValue("@slotCount")
		Set rs = sphn.rs
		Set rs2 = sphn.rs2
		Set rs3 = sphn.rs3
		Set rs4 = sphn.rs4
	End If
	set sphn = Nothing

	retString = slotCount & SEPARATOR & accountName	
	
	i=0
	retString = retString & SEPARATOR
	Do Until rs.Eof
		If i>0 Then
			retString = retString & SUBSEPARATOR
		End If
		
		retString = retString & rs(0)
		
		rs.MoveNext  : i=i+1
	Loop
	
	Do Until rs2.Eof
		retString = retString & SEPARATOR & rs2(0) &_
				SEPARATOR & rs2(1) & SEPARATOR & rs2(2) &_
				SEPARATOR & rs2(3) & SEPARATOR & rs2(4) &_
				SEPARATOR & rs2(5) & SEPARATOR & rs2(6) &_
				SEPARATOR & rs2(7) & SEPARATOR & rs2(8) &_
				SEPARATOR & rs2(9) & SEPARATOR & FormatDt(rs2(10),"SQL_TM") &_
				SEPARATOR & rs2(11) & SEPARATOR & rs2(12) &_
				SEPARATOR & rs2(13)

		rs3.Filter = "characterName='"& rs2(0) &"'"
		i=0
		If rs3.Eof Then retString = retString & SEPARATOR
		Do Until rs3.Eof
			If i=0 Then
				retString = retString & SEPARATOR
			Else
				retString = retString & SUBSEPARATOR
			End If
			retString = retString & rs3(1) & SUBSEPARATOR & rs3(2)
			rs3.MoveNext : i=i+1
		Loop
		rs3.Filter = adFilternone
		
		rs4.Filter = "characterName='"& rs2(0) &"'"
		i=0
		If rs4.Eof Then retString = retString & SEPARATOR
		Do Until rs4.Eof
			If i=0 Then
				retString = retString & SEPARATOR
			Else
				retString = retString & SUBSEPARATOR
			End If
			retString = retString & rs4(1) & SUBSEPARATOR & rs4(2) & SUBSEPARATOR & rs4(3) & SUBSEPARATOR & rs4(4) & SUBSEPARATOR & FormatDt(rs4(5),"SQL_TM")
			rs4.MoveNext : i=i+1
		Loop
		rs4.Filter = adFilternone

		rs2.MoveNext
	Loop
	rs4.Close
	rs3.Close
	rs2.Close
	rs.Close
	
	Call OK(retString)
%>