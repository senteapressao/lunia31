<%@ Language=VBScript %>
<%Option Explicit%>
<%
	' -- common codes ------------------------------------------------------
	Dim characterDBconnectionString,stageDBconnectionString,guildDBconnectionString,gateDBconnectionString,cashDBconnectionString
	characterDBconnectionString	= GetConnStr("DESKTOP-S29QIQ5\SQLEXPRESS","v3_character","sa","algoraSaturn2203@")
	stageDBconnectionString		= GetConnStr("DESKTOP-S29QIQ5\SQLEXPRESS","v4_stage","sa","algoraSaturn2203@")
	guildDBconnectionString		= GetConnStr("DESKTOP-S29QIQ5\SQLEXPRESS","v3_guild","sa","algoraSaturn2203@")
	gateDBconnectionString		= GetConnStr("DESKTOP-S29QIQ5\SQLEXPRESS","v3_gate","sa","algoraSaturn2203@")
	cashDBconnectionString		= GetConnStr("DESKTOP-S29QIQ5\SQLEXPRESS","d-shop","sa","algoraSaturn2203@")
	
	' predefined constant
	const ServerGroupName = "TW-R"
	const ServerGroupCode = 126
	const LoginWSDL ="http://DESKTOP-S29QIQ5\SQLEXPRESS/ServiceLogin.asmx?wsdl"
	
	Dim remoteIp
	remoteIp=Request.ServerVariables("REMOTE_ADDR")
	
	' predefined(reserved) variables
	Dim SEPARATOR, SUBSEPARATOR, OPERATION, MINIMALPARAMETER, REQUESTERID
	Dim Parameters

	' // functions ////////////////////////////////////////////////////////////////////////////
	Function GetConnStr(SERVER,DBNAME,UID,PWD)
		'GetConnStr = "DRIVER=SQL Server;SERVER="& SERVER &";DATABASE="& DBNAME &";Network=DBMSSOCN;Address="& SERVER &";User Id="& UID &";PASSWORD="& PWD &";"	
		GetConnStr = "Provider=SQLOLEDB;Data Source="& SERVER &";Initial Catalog="& DBNAME &";user ID="& UID &";password="& PWD &";"
	End Function
	
	' USE this function - ExecSP like this
	'		Dim ret, params
	'		params=Array("1", "2")
	'		ret=ExecSP(connectionString, "test1", params)
	'		' params gets array if the stored procedure has output parameters
	Function ExecSP(connectionString, procedure, params)
		Dim connection, command, ret
		
		Set connection=Server.CreateObject("ADODB.Connection")
		Set command=Server.CreateObject("ADODB.Command")			
		
		connection.Open(connectionString)
		command.ActiveConnection=connection
		command.CommandType = 4 ' adCmdStoredProc
		command.CommandText = procedure	
		
		If IsArray(params) Then
			Dim i
			For i=0 To UBound(params)
				command.Parameters(i+1).Value=params(i)
			Next
			command.Execute()
			ExecSP=command.Parameters(0) ' return value
		Else ' to make non-array support
			If params<>"" Then command.Parameters(1)=params
			command.Execute()
			ExecSP=command.Parameters(0) ' return value
		End If
		
		Dim retcnt
		retcnt=0
		For i=1 To command.Parameters.Count-1
			If command.Parameters(i).Direction=3 Then retcnt=retcnt+1 ' number of output paramters 
		Next
		If retcnt=0 Then params="" Else Redim params(retcnt-1) ' set return parameters
		retcnt=0
		For i=1 To command.Parameters.Count-1
			If command.Parameters(i).Direction=3 Then ' OUTPUT
				params(retcnt)=command.Parameters(i).Value
				retcnt=retcnt+1
			End If
		Next
		
		connection.Close()
		
		Set command=Nothing
		Set connection=Nothing
	End Function
					
	' return parameters
	'		ex) Call Ok(Array("aaaa", "bbbb", "cccc"))
	Function Ok(params)
		'Response.Write(RequestID & SEPARATOR & OPERATION)
		If IsArray(params) Then
			Dim i
			For i=0 To UBound(params)
				Response.Write(params(i))
				If i<>UBound(params) Then Response.Write(SEPARATOR)
			Next
		Else
			If Not IsNull(params) Then
				Response.Write(params)
			End If
		End If
		
		Response.End()
	End Function
	
	Function Error(errorCode)
		If IsNumeric(errorCode) Then
			Response.Status=600+errorCode
			Response.Write("ERROR NUMBER :" & errorCode)
		Else
			Response.Status=699
			Response.Write("ERROR :" & errorCode)
		End If
		Response.End()
	End Function
	
	Function DebugError(errorMessage)
		Response.Status=699
		Response.Write("ERROR : " & errorMessage)
		Response.End()
	End Function

	Function SqlQuot(str)
		SqlQuot="N'" & Replace(str, "'", "''") & "'"
	End Function
	
	Function URLDecode(byVal encodedstring)
		Dim strIn, strOut, intPos, strLeft
		Dim strRight, intLoop
		strIn  = encodedstring : strOut = "" : intPos = Instr(strIn, "+")
		Do While intPos
			strLeft = "" : strRight = ""
			If intPos > 1 then strLeft = Left(strIn, intPos - 1)
			If intPos < len(strIn) then strRight = Mid(strIn, intPos + 1)
			strIn = strLeft & " " & strRight
			intPos = InStr(strIn, "+")
			intLoop = intLoop + 1
		Loop
		intPos = InStr(strIn, "%")
		Do while intPos
			If intPos > 1 then strOut = strOut & Left(strIn, intPos - 1)
			strOut = strOut & Chr(CInt("&H" & mid(strIn, intPos + 1, 2)))
			If intPos > (len(strIn) - 3) then
				strIn = ""
			Else
				strIn = Mid(strIn, intPos + 3)
			End If
			intPos = InStr(strIn, "%")
		Loop
		URLDecode = strOut & strIn
	End Function		
	
	' internal initialize without parameter validation
	' Op : Request operation
	Function InitL(Op)
		SUBSEPARATOR=Chr(11)
		SEPARATOR=Chr(8)
		OPERATION=Op
		REQUESTERID=Request.ServerVariables("HTTP_REQUESTER") ' "HTTP_" prefix is attached by IIS
		If REQUESTERID="" Then REQUESTERID="(empty)"
		
		Parameters=Split(URLDecode(Request.QueryString), SEPARATOR)
	End Function
	dim parametro, nomeserver, separador, especial
	
	Function InitD(parametro,nomeserver,separador)
		Dim Url, Op
		Url=Request.ServerVariables("URL")
		Url=Split(Url, ".")(0)
		
		Op=Mid(Url, InStrRev(Url, "/")+1)
		SUBSEPARATOR=Chr(11)
		SEPARATOR=Chr(8)
		especial=Chr(35)
		OPERATION=Op
		REQUESTERID=nomeserver ' "HTTP_" prefix is attached by IIS
		if separador =true Then entrada = Replace(parametro,especial,"")
		If REQUESTERID="" Then REQUESTERID="(empty)"
		
		Parameters=Split(URLDecode(parametro), SEPARATOR)
		
	End Function
	
	' basic initialize
	Function Init()
		Dim Url, Op
		Url=Request.ServerVariables("URL")
		Url=Split(Url, ".")(0)
		
		Op=Mid(Url, InStrRev(Url, "/")+1)
		InitL(Op)
	End Function

	' strict initializing and preparing service
	Function InitS(Op, Mp)
		MINIMALPARAMETER=Mp
		Call InitL(Op)
		If UBound(Parameters)<(MINIMALPARAMETER-1) Then
			call Error (1)
		End If
	End Function

	Function FormatDt(argDt,argType)
		Select Case CStr("" & argType)
			Case "ISO"
				FormatDt = Year(argDt) & Right("0" & Month(argDt), 2) & Right("0" & Day(argDt),2)
			Case "KOR"
				FormatDt = Year(argDt) & "/" & Right("0" & Month(argDt), 2) & "/" & Right("0" & Day(argDt), 2)
			Case "ENG"
				FormatDt = Right("0" & Month(argDt), 2) & "/" & Right("0" & Day(argDt),2) & "/" & Year(argDt)
			Case "ISO_TM"
				FormatDt = Year(argDt) & Right("0" & Month(argDt), 2) & Right("0" & Day(argDt), 2)
				FormatDt = FormatDt & Right("0" & Hour(argDt), 2) & Right("0" & Minute(argDt), 2) & Right("0" & Second(argDt), 2)
			Case "KOR_TM"
				FormatDt = Year(argDt) & "/" & Right("0" & Month(argDt), 2) & "/" & Right("0" & Day(argDt), 2) & ", "
				FormatDt = FormatDt & Right("0" & Hour(argDt), 2) & ":" & Right("0" & Minute(argDt), 2) & ":" & Right("0" & Second(argDt), 2)
			Case "ENG_TM"
				FormatDt = Right("0" & Month(argDt), 2) & "/" & Right("0" & Day(argDt),2) & "/" & Year(argDt) & ", "
				FormatDt = FormatDt & Right("0" & Hour(argDt), 2) & ":" & Right("0" & Minute(argDt), 2) & ":" & Right("0" & Second(argDt), 2)
			Case "SQL"
				FormatDt = Year(argDt) & "-" & Right("0" & Month(argDt), 2) & "-" & Right("0" & Day(argDt), 2)
			Case "SQL_TM"
				FormatDt = Year(argDt) & "-" & Right("0" & Month(argDt), 2) & "-" & Right("0" & Day(argDt),2) & " "
				FormatDt = FormatDt & Right("0" & Hour(argDt), 2) & ":" & Right("0" & Minute(argDt), 2) & ":" & Right("0" & Second(argDt), 2)
			Case "SQL_WK"
				FormatDt = Year(argDt) & "-" & Right("0" & Month(argDt), 2) & "-" & Right("0" & Day(argDt), 2) & " (" & WeekDayName(WeekDay(LogDate)) & ")"
			Case "SSN"
				FormatDt = Right("0" & Year(argDt), 2) & Right("0" & Month(argDt), 2) & Right("0" & Day(argDt), 2)
			Case Else
				FormatDt = ""
		End Select
	End Function
	
	Function getConvAscii(value)
		Dim strV,i,s,intChr,ichr1,ichr2,strS
		 s = LenB(value) : strV = "" :
		 For i = 1 to s
			intChr = AscB(MidB(value,i,2))
			if intChr > 128 then
				ichr1 = intChr
				i = i + 1
				ichr2 = AscB(MidB(value,i,2))
				strS = "&H" & Hex(ichr1) & Hex(ichr2)
				strV = strV & Chr(strS)
			Else
				strV = strV & Chr(intChr)
			End If
		 Next
		 getConvAscii = strV
	End Function
	
	Function GetAccountName(characterName)
		Dim sphn,blsResult : blsResult = False
		Dim accountName : accountName = ""
		
		Set sphn = new SPHelper_NoTran
		with sphn
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = characterDBconnectionString
			.SPName = "dbo.GetAccountName"
			Call .InitCommand()
			Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
			Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
			Call .AppendParam("@accountName",adVarWChar,adParamOutput,50,null)
			blsResult = .ExecNoRecords()
		End with
		
		If blsResult Then
			Dim ret
			ret = sphn.GetParamValue("RETURN_VALUE")
			if ret<>0 Then
				Call Error(ret)
			End If
			accountName = sphn.GetParamValue("@accountName")
		End If
		set sphn = Nothing
		
		GetAccountName = accountName
	End Function
	
	' //////////////////////////////////////////////////////////////////////////// functions //
%>