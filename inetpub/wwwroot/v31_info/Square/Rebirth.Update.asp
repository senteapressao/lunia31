<!--#include file="../common.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<%
	Call Init()
	' request  : (POST)
	'			characterName,stageLevel,stageExp,pvpLevel,pvpExp,warLevel,warExp,skillPoint,rebirthCount,storedLevel,storedSkillPoint,lastRebirthDate
	'			, initLicenses(bool)
	'			, delete Items(subseparated: bagNumber,positionNumber)
	'			, update Items(subseparated: itemHash,bagNumber,positionNumber,stackedCount,itemSerial,instance,itemExpire)
	'			, delete Skills(subseparated: skillGroupHash)
	'			, update Skills(subseparated: skillGroupHash,skillLevel)
	' response : characterName
	
	Dim i,j,k
	Dim sql,conn,rs
	Dim retString : retString = ""
	
	Dim rowData
	rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)

	' preparing request parameters
	Dim characterName,stageLevel,stageExp,pvpLevel,pvpExp,warLevel,warExp,skillPoint _
		,rebirthCount,storedLevel,storedSkillPoint,lastRebirthDate
	Dim initLicenses
	Dim deleteItems,updateItems,insertItems _
		,deleteSkills,updateSkills,insertSkills
	Dim arr

	If UBound(Parameters)<16 Then Call Error("not enough parameter : ("& UBound(Parameters) &");"& rowData)
	
	characterName	= Parameters(0)
	stageLevel		= Parameters(1)
	stageExp		= Parameters(2)
	pvpLevel		= Parameters(3)
	pvpExp			= Parameters(4)
	warLevel		= Parameters(5)
	warExp			= Parameters(6)	
	skillPoint		= Parameters(7)
	
	rebirthCount	= Parameters(8)
	storedLevel		= Parameters(9)
	storedSkillPoint= Parameters(10)
	lastRebirthDate	= Parameters(11)
	
	initLicenses	= Parameters(12)

	deleteItems		= Parameters(13)
	updateItems		= Parameters(14)
	deleteSkills	= Parameters(15)
	updateSkills	= Parameters(16)
	
	Dim setCount : setCount = 0

	' update basic character
	sql = ""
	sql = sql &"set lock_timeout 2000"& vbCrLf
	'sql = sql &"set xact_abort on"& vbCrLf
	sql = sql &"begin tran"& vbCrLf
	sql = sql &"declare @nError int, @strSP varchar(50)"& vbCrLf
	
	' basic info
	sql = sql &"exec Rebirth_Update"&_
		" "& SqlQuot(characterName) &_
		","& stageLevel &_
		","& stageExp &_
		","& pvpLevel &_
		","& pvpExp &_
		","& warLevel &_
		","& warExp &_
		","& skillPoint &_
		","& rebirthCount &_
		","& storedLevel &_
		","& storedSkillPoint &_
		","& SqlQuot(FormatDt(lastRebirthDate,"SQL_TM")) &_
		",@nError output, @strSP output"&_
		vbCrLf
	setCount = setCount +1
	sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
	
	' init Licenses
	If CInt(initLicenses)=1 Then
		sql = sql &"exec Rebirth_InitLicenses"&_
			" "& SqlQuot(characterName) &_
			",@nError output, @strSP output"&_
			vbCrLf
		setCount = setCount +1
		sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
	End If

	' delete Items
	arr = Split(deleteItems,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 2)<>0 Then
			Call Error("invalid number of item parameter")
		End If
		
		For i=0 To UBound(arr) Step 2
			sql = sql &"exec DeleteItem"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				","& arr(i+1) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If
	
	' update Items
	arr = Split(updateItems,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 7)<>0 Then
			Call Error("invalid number of item parameter")
		End If
		
		For i=0 To UBound(arr) Step 7			
			sql = sql &"exec UpdateItem"&_
				" "& SqlQuot(characterName) &_
				","& arr(i+1) &_
				","& arr(i+2) &_
				","& arr(i) &_
				","& arr(i+3)
			If IsNumeric(arr(i+4)) Then
				sql = sql &","& arr(i+4)
			Else
				sql = sql &",null"
			End If				
			sql = sql &","& arr(i+5) &_
				","& SqlQuot(arr(i+6)) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If
	
	' delete Skills
	arr = Split(deleteSkills,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		For i=0 To UBound(arr) Step 1
			sql = sql &"exec DeleteSkill"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If
	
	' update Skills
	arr = Split(updateSkills,SUBSEPARATOR)
	If Ubound(arr)>=0 Then
		If ((UBound(arr)+1) Mod 2)<>0 Then
			Call Error("invalid number of item parameter")
		End If
		
		For i=0 To UBound(arr) Step 2
			sql = sql &"exec UpdateSkill"&_
				" "& SqlQuot(characterName) &_
				","& arr(i) &_
				","& arr(i+1) &_
				",@nError output, @strSP output"&_
				vbCrLf
			setCount = setCount +1
			sql = sql & "if(@nError not in (0)) goto lblFail"& vbCrLf
		Next
	End If
	
	sql = sql & "lblSuccess:"& vbCrLf
	sql = sql & "	commit tran"& vbCrLf
	'sql = sql & "	insert	dbo.tblError(nError,strSP) select @nError,@strSP "& vbCrLf
	sql = sql & "	goto lblEnd"& vbCrLf
	sql = sql & "lblFail:"& vbCrLf
	sql = sql & "	rollback tran"& vbCrLf
	sql = sql & "	insert	dbo.tblError(characterName,strASP,nError,strSP) select "& SqlQuot(characterName) &","& SqlQuot(Request.ServerVariables("URL")) &",@nError,@strSP "& vbCrLf
	sql = sql & "	goto lblEnd"& vbCrLf
	sql = sql & "lblEnd:"& vbCrLf
	'sql = sql &"commit tran"& vbCrLf
	
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open(characterDBconnectionString)
	'conn.Execute(sql)
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.Open sql,conn,adOpenStatic,adLockReadOnly,adCmdText
	For i=1 To setCount
		set rs = rs.NextRecordset
	Next
	conn.Close()
	
	Call Ok(characterName)
%>