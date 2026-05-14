<%
	Dim SEPARATOR,SUBSEPARATOR
	Dim Parameters
	SUBSEPARATOR=Chr(11)
	SEPARATOR=Chr(8)
	Dim temp
	temp = Replace(Request.QueryString,"%08",Chr(8))
	temp = Replace(temp,"%11",Chr(11))
	Parameters = Split(temp,SEPARATOR)
	
	Function Logger()
        Dim arrPath, fileName, reciver, dir, tfile, count, objFSO, data
        arrPath = Split(Request.ServerVariables("SCRIPT_NAME"), "/") '/
        
        Dim category 
        category = arrPath(UBound(arrPath)-1)
        
        fileName = arrPath(UBound(arrPath))
        fileName= Split(fileName,".asp")(0)

		Logger = getConvAscii(Request.BinaryRead(Request.TotalBytes))

        if fileName <> "ListSquareStatus" and fileName <> "Ping.Check" then
            dir = Server.MapPath("/logs") & "/[" & category & "]"&fileName&".txt"
            Set objFSO = CreateObject("Scripting.FileSystemObject")
            Set tfile=objFSO.OpenTextFile(dir, 8, true)
            tfile.WriteLine("Time: " & Now & " - "&fileName&" : " & Request.QueryString)
            tfile.WriteLine("REQUESTERID : " & Request.ServerVariables("HTTP_REQUESTER"))
            tfile.WriteLine("rowData : " & Logger)
            tfile.WriteLine("-------------------------------------------------------------------------------------------------------------")
            tfile.close
			set tfile = nothing
			set objFSO = nothing
        
        end if
    End Function
	
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
		Response.End
	End Function
	
	Function Error(errorCode)
		If IsNumeric(errorCode) Then
			Response.Write("ERROR NUMBER :" & errorCode)
		Else
			Response.Write("ERROR :" & errorCode)
		End If
		Response.End
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
				strV = strV & ChrW(strS)
			Else
				strV = strV & ChrW(intChr)
			End If
		 Next
		 getConvAscii = strV
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
%>