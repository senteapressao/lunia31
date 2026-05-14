<%
	Dim arrServerCode
	Set arrServerCode = CreateObject("Scripting.Dictionary")
	
	arrServerCode.ADD 127,"test"
	
	' Response.Write
	Sub Print(str)
		Response.Write str
	End Sub
	
	Sub Println(str)
		Response.Write str &"<br />"
	End Sub

	'********************************************************************
	'********************************************************************	
	Function Decrypt(encText)
		Dim oEnc,password
		password = "Guilds_Ranki"
		set oEnc = Server.CreateObject("DynuEncrypt.Functions")		
		
		Decrypt = oEnc.Decrypt(encText,password)
		
		set oEnc = Nothing		
	End Function
	
	Function Quot(ByVal Value)
		Quot = "'" & Value & "'"
	End Function
	
	Sub ChkInjection(ByVal str)
		If InStr(LCase(str),"unicode")>0 OR _
		InStr(LCase(str)," between")>0 OR _
		InStr(LCase(str),"substring")>0 OR _
		InStr(LCase(str)," and")>0 OR _
		InStr(LCase(str)," or")>0 OR _
		InStr(LCase(str)," having")>0 OR _
		InStr(LCase(str)," group")>0 OR _
		InStr(LCase(str)," openrowset")>0 OR _
		InStr(LCase(str)," opendatasource")>0 OR _
		InStr(LCase(str)," cmdshell")>0 OR _
		InStr(LCase(str)," select")>0 OR _
		InStr(LCase(str)," update")>0 OR _
		InStr(LCase(str)," insert")>0 OR _
		InStr(LCase(str)," delete")>0 OR _
		InStr(LCase(str),"'")>0 OR _
		InStr(LCase(str),"""")>0 OR _
		InStr(LCase(str),"--")>0 OR _
		InStr(LCase(str),";")>0 Then
			Call ThrowError(null,"illegal access:0","close")
		End If
	End Sub

	Function FiltSQL(str)
		FiltSQL = Replace(str, "'", "''")		 
	End Function 

	Function FiltHTML(str)
		str = Replace(str, "&", "&amp;")
		str = Replace(str,"<","&lt;")
		str = Replace(str,">","&gt;")
		str = Replace(str, chr(39), "&#39") 
		str = Replace(str, chr(34), "&#34")
		FiltHTML = str
	End Function 

	Function FiltBR(str)
		str = Replace(str, vbcrlf, "<br />" )
		FiltBR = str
	End Function 

	Function FiltScript(str)
		str = Replace(str,"<script","<x-script")
		str = Replace(str,"</script","</x-script")
		FiltScript = str
	End Function

	' FiltFlash
	Function FiltFlash(str)
		str = Replace(str,"<","＜")
		str = Replace(str,">","＞")
		FiltFlash = str
	End Function 

	Function FiltSearch(str,strSearch,color)
		FiltSearch = Replace(str,strSearch,"<FONT color="""& color &""">"& strSearch &"</FONT>")		
	End Function 

	Function CutString(str,strlen)
		Dim rValue
		Dim nLength
		Dim strF,tmpStr,tmpLen
		nLength = 0.00
		rValue = ""

		For strF = 1 to Len(str)
			tmpStr = MID(str,strF,1)
			tmpLen = ASC(tmpStr)
			If  (tmpLen < 0) Then
				nLength = nLength + 1.2
				rValue = rValue & tmpStr
			ElseIf (tmpLen >= 97 And tmpLen <= 122) Then
				nLength = nLength + 0.63
				rValue = rValue & tmpStr
			ElseIf (tmpLen >= 65 And tmpLen <= 90) Then 
				nLength = nLength + 0.81
				rValue = rValue & tmpStr
			ElseIf (tmpLen = 32) Then 
				nLength = nLength + 0.5
				rValue = rValue & tmpStr
			Else
				nLength = nLength + 0.9
				rValue = rValue & tmpStr			
			End If
			If (nLength > strlen) Then
				rValue = rValue & ".."
				Exit For
			End If
		Next 
		CutString = rValue
	End Function

	Function ChgDec(str)
		ChgDec = Right("00"&str,2)
	End Function
	
	'********************************************************************
	'********************************************************************
	Function MakeSearchQuery(SearchQuery,Str)
		If SearchQuery="" Then
			SearchQuery = " WHERE "& Str
		Else
			SearchQuery = SearchQuery &" AND "& Str
		End If
		MakeSearchQuery = SearchQuery
	End Function
	
	Function GetGuid()
		Dim objTypeLib, strGuid
		Set objTypeLib = Server.CreateObject("Scriptlet.Typelib")
			strGuid = objTypeLib.Guid
		Set objTypeLib = Nothing
		strGuid = Replace(strGuid,"-","")
		strGuid = Replace(strGuid, "{", "")
		strGuid = Replace(strGuid, "}", "")
		GetGuid = strGuid
		Exit Function
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
	
	' convert date
	Function FormatDate(tDate,DateType)
		If Not(IsDate(tDate)) Then FormatDate=tDate:Exit Function
		Dim retVal
		Dim tmpDate
		Dim tYear,tMonth,tDay,tWeekDay
		tYear	= Year(tDate)
		tMonth	= Month(tDate)
		tDay	= Day(tDate)
		Select Case WeekDay(tDate)
			Case 1
				tWeekDay = "Sunday"
			Case 2
				tWeekDay = "Monday"
			Case 3
				tWeekDay = "Tuesday"
			Case 4
				tWeekDay = "Wednesday "
			Case 5
				tWeekDay = "Thursday"
			Case 6
				tWeekDay = "Friday"
			Case 7
				tWeekDay = "Saturday"
		End Select	
		
		If Hour(writedate)>12 Then
			tmpDate = chgDec(Hour(tDate)-12) &":"& chgDec(minute(tDate))&" PM"
		Else 
			tmpDate = chgDec(Hour(tDate)) &":"& chgDec(minute(tDate))&" AM"
		End IF

		If DateType=1 Then		
			retVal = tWeekDay
		ElseIf DateType=2 Then
			If DateDiff("d",Now(),tDate)<0 Then 'prev
				retVal=ChgDec(Month(tDate)) &"-"& ChgDec(Day(tDate))
			Else
				retVal=tmpDate
			End If
		ElseIf DateType=3 Then
			retVal = chgDec(Month(tDate)) &"-"& chgDec(day(tDate)) &"-"& Year(tDate) &" "& tmpDate
		ElseIf DateType=4 Then
			retVal = DateMonthName(tDate)&" "& chgDec(day(tDate)) &","& Year(tDate)
		End If
		FormatDate = retVal
	End Function
	
	Function DateMonthName(value)
		Dim retStr
		Select Case Month(value)
			Case 1
				retStr = "January"
			Case 2
				retStr = "February"
			Case 3
				retStr = "March"
			Case 4
				retStr = "April"
			Case 5
				retStr = "May"
			Case 6
				retStr = "June"
			Case 7
				retStr = "July"
			Case 8
				retStr = "August"
			Case 9
				retStr = "September"
			Case 10
				retStr = "October"
			Case 11
				retStr = "November"
			Case 12
				retStr = "December"
		End Select	
		DateMonthName=retStr
	End Function
	

	Function GetPaging(pageSize,pagingCnt,currentPage,totalCnt,listURL,pagingQuery)
		Dim Result,PageNo
		Dim sPage,ePage,lPage
		
		If pageSize=0 Then pageSize=1
		If pagingCnt=0 Then pagingCnt=1
		
		PageNo = 0
		sPage = int((currentPage-1)/pagingCnt)*pagingCnt+1
		ePage = sPage+pagingCnt-1
		lPage = int((totalCnt-1)/pageSize)+1
		
		Result = ""
		
		Result = Result &	"<table cellpadding='0' cellspacing='0'>"&_
							"<tr height=13>"&_
							"	<td><img src="""& ImgURL &"/board/btn_first.gif"" class=""vam""/></td>"&_
							"	<td>&nbsp;&nbsp;"
		If sPage>1 Then 
			Result = Result &"<a href='"& listURL &"?GotoPage="& sPage-1 &"&"& pagingQuery &"'><img src="""& ImgURL &"/board/btn_prev.gif"" class=""vam""/></a>"
		Else
			Result = Result &"<img src="""& ImgURL &"/board/btn_prev.gif"" class=""vam""/> "
		End If

		For PageNo=sPage To ePage
			If PageNo = Cint(CurrentPage) Then 
				Result = Result &" <B>"& pageNo &"</B> "
			Else 
				Result = Result &" <A href='"& listURL &"?GotoPage="& pageNo &"&"& pagingQuery &"'>"& pageNo &"</A> " 
			End If
			If PageNo>=lPage Then Exit For
		Next
		
		If ePage<lPage  Then 
			Result = Result &"<a href='"& listURL &"?GotoPage="& ePage+1 &"&"& pagingQuery &"'>"& " <img src="""& ImgURL &"/board/btn_next.gif"" class=""vam""/></a>" 
		Else
			Result = Result &"<img src="""& ImgURL &"/board/btn_next.gif"" class=""vam""/>"
		End If
		
		If lPage=0 Then lPage=1
		Result = Result &	"	</td>"&_
							"	<td>&nbsp;&nbsp;<img src="""& ImgURL &"/board/btn_last.gif"" class=""vam""/></td>"&_
							"</tr>"&_
							"</table>"
		GetPaging = Result
	End Function
	
	Function GetPagingShop(pageSize,pagingCnt,currentPage,totalCnt,listURL,pagingQuery)
		Dim Result,PageNo
		Dim sPage,ePage,lPage
		
		If pageSize=0 Then pageSize=1
		If pagingCnt=0 Then pagingCnt=1
		
		PageNo = 0
		sPage = int((currentPage-1)/pagingCnt)*pagingCnt+1
		ePage = sPage+pagingCnt-1
		lPage = int((totalCnt-1)/pageSize)+1
		
		Result = ""
		If sPage>1 Then 
			Result = Result &"<a href='"& listURL &"?GotoPage="& sPage-1 &"&"& pagingQuery &"'><img src="""& ImgURL &"/shop_v1/btn/btn_arrow_left.gif"" class=""vam""/></a>"
		Else
			Result = Result &"<img src="""& ImgURL &"/shop_v1/btn/btn_arrow_left.gif"" class=""vam""/> "
		End If

		For PageNo=sPage To ePage
			If PageNo = Cint(CurrentPage) Then 
				Result = Result &" <span class='fc1 b'>"& pageNo &"</span> "
			Else 
				Result = Result &" <a href='"& listURL &"?GotoPage="& pageNo &"&"& pagingQuery &"'>"& pageNo &"</a> " 
			End If
			If PageNo>=lPage Then Exit For
		Next
		
		If ePage<lPage  Then 
			Result = Result &"<a href='"& listURL &"?GotoPage="& ePage+1 &"&"& pagingQuery &"'>"& " <img src="""& ImgURL &"/shop_v1/btn/btn_arrow_right.gif"" class=""vam""/></a>" 
		Else
			Result = Result &"<img src="""& ImgURL &"/shop_v1/btn/btn_arrow_right.gif"" class=""vam""/>"
		End If
		
		GetPagingShop = Result
	End Function
	
	' Search
	Function GetSearch(SearchField,SearchText,listURL,searchQuery,field)
		If Not(IsObject(field)) Then
			Set field = CreateObject("Scripting.Dictionary")
			field.ADD "Subject","subject"
			field.ADD "Content","content"
			field.ADD "WriterID","WriterID"
			field.ADD "WriterName","WriterName"
		End If
		
		Dim Result,key
		Result = ""
		Result = Result &"<script type=""text/javascript"">"& vbCrLf
		Result = Result &"function validateSearch() {"& vbCrLf
		Result = Result &"var frmSearch=document.getElementById(""frmSearch"");"& vbCrLf
		Result = Result &"frmSearch.SearchText.value=frmSearch.SearchText.value.replace(""#"","""");"& vbCrLf
		Result = Result &"if(frmSearch.SearchText.value=="""") {alert('Digite a palavra de busca.');return false;}"& vbCrLf
		Result = Result &"location.href="""& listURL &"?"
		If searchQuery<>"" Then  Result = Result & searchQuery &"&"
		Result = Result &"SearchField=""+frmSearch.SearchField.value+""&SearchText=""+frmSearch.SearchText.value;}"& vbCrLf
		Result = Result &"</script>" & vbCrLf
		Result = Result &"<form method=""get"" id=""frmSearch"" onsubmit=""validateSearch();return false;"">"& vbCrLf
		Result = Result &"<table cellpadding=""0"" cellspacing=""0"">"& vbCrLf
		Result = Result &"<tr><td>"& vbCrLf
		Result = Result &"<select name=""SearchField"" style=""width:100px;"">"& vbCrLf
		
		For Each key In field
			Result = Result &"<option value="""& key &"""> "& field(key) &" </option>"& vbCrLf
		Next
		
		Result = Result &"</select>"& vbCrLf
		Result = Result &"<script type=""text/javascript"">document.getElementById(""frmSearch"").SearchField.value='"& SearchField &"';</script></td>"& vbCrLf
		Result = Result &"<td>&nbsp;&nbsp;<input type=""text"" name=""SearchText"" maxlength=""20"" value="""& SearchText &""" class=""input_t vam"" style=""width:100px;"" /></td>"& vbCrLf
		Result = Result &"<td>&nbsp;&nbsp;<input type=""image"" src="""& ImgURL &"/board/btn_search.gif"" class=""vam"" /></td>"& vbCrLf
		Result = Result &"<td>"& vbCrLf
		If SearchText<>"" Then 
			Result = Result &"&nbsp;<a href=""javascript:location.href='"& listURL &"?"& searchQuery &"'""><img src="""& ImgURL &"/board/btn_view.gif"" class=""vam""   style=""vertical-align:middle;""/></a>"& vbCrLf
		End If		
		Result = Result &"</td></tr></table></form>"& vbCrLf
		
		GetSearch = Result
	End Function
	
	'********************************************************************
	'********************************************************************

	Function GetFileLen(size)
		If size<1024 Then
			GetFileLen = FormatNumber(size,0)&"Byte"
		ElseIf size/1024<1024 Then
			GetFileLen = FormatNumber(size/1024,2)&"KB"
		ElseIf size/1024>=1024 Then
			GetFileLen = FormatNumber(size/1024/1024,2)&"MB"
		End If
	End Function
	
	' return file extend Image
	Function GetFileTypeImg(fileType)
		Select Case Ucase(fileType)
			Case "HWP", "COM", "EXE","TXT", "WAV", "XLS", "DOC", "RAR", "ACE", "GIF", "BMP", "MP3", "PDF", "HLP", "CHM","PPT"
				GetFileTypeImg = fileType &".gif"
			Case "ZIP", "Z", "GZ", "TAR"
				GetFileTypeImg = "zip.gif"
			Case "RA", "RAM"
				GetFileTypeImg = "ra.gif"
			Case "HTM", "HTML"
				GetFileTypeImg = "htm.gif"
			Case "JPG", "JPEG"
				GetFileTypeImg = "jpg.gif"
			Case "MPG", "MPEG", "AVI", "ASF"
				GetFileTypeImg = "mpg.gif"
			Case else
				GetFileTypeImg = "etc.gif"
		End Select
	End Function

	Function ChkFileType(fileType,usefulType)
		Dim i,arrUsefulType
		arrUsefulType = Split(usefulType,",")
		For i=0 To Ubound(arrUsefulType)
			If Ucase(fileType)=UCase(arrUsefulType(i)) Then
				ChkFileType = True : Exit Function
			End If
		Next
		ChkFileType = False
		
	End Function
	
	Function ChkFileLen(fileLen,maxLen)
		If fileLen>maxLen Then
			ChkFileLen = False : Exit Function
		End If
		ChkFileLen = True
	End Function
	
	Function GetUniqueFileName(fileLoc,fileName)
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject") ' FileSystemObject 를 연다.

		Dim tmpFileName,tmpFileExt,fileCheck,tmpFileName2,tmpFileName3
		Dim FileNo : FileNo=1
		
		tmpFileName =  Left(fileName,InStrRev(fileName,".")-1)
		tmpFileExt  = Mid(fileName,InStrRev(fileName,".")+1)

		fileCheck = false
		Do
			If(fso.FileExists(fileLoc & tmpFileName &"."& tmpFileExt)) Then
				If InStr(tmpFileName,"(")>0 AND InStr(tmpFileName,")")>0 Then
					tmpFileName2 = Left(tmpFileName,InStrRev(tmpFileName,"(")-1)
					tmpFileName3 = Mid(tmpFileName,InStrRev(tmpFileName,"("))
					FileNo = Replace(Replace(tmpFileName3,"(",""),")","")
					tmpFileName = tmpFileName2 &"("& FileNo+1 &")"
				Else
					tmpFileName = tmpFileName &"("& FileNo &")"
				End If
				fileCheck = false
			Else
				fileCheck = true
			End If
			FileNo = FileNo+1
		Loop Until  fileCheck
		GetUniqueFileName = tmpFileName&"."&tmpFileExt
	End Function

	' save file
	Function UploadFile(uploadField,fileLoc,newfileName)
		Dim FileName
		FileName = GetUniqueFileName(fileLoc,newfileName)
		uploadField.SaveAs fileLoc & FileName
	End Function
	
	Function LuniPointPage(page)
		Dim arrPage
		arrPage = Array("ifrm_charge.asp","ifrm_charging_log.asp","ifrm_shopping_log.asp")
		LuniPointPage = arrPage(Cint(Page))
	End Function
	
	Sub PrintRS(RS)
		Dim i,j,k
		If TypeName(RS)="Recordset" Then
			Do Until RS.Eof
				If i=0 Then
					Print "<table border='1'>"
					Print "<tr>"
					For Each j In RS.Fields
						Print "<td style='font-size:11px;'>"
						Print j.Name
						Print "</td>"
					Next
					Print "</tr>"
				End If
				Print "<tr height='16'>"
				For Each j In RS.Fields
					Print "<td style='font-size:11px;'>"
					Print RS(j.Name)
					Print "</td>"
				Next
				Print "</tr>"
				RS.MoveNext : i=i+1
				If RS.Eof Then Print "</table>"
			Loop
		End If
	End Sub
%>