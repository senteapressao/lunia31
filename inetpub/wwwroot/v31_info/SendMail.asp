<%
	Function SendMail(characterName,sender,title,message,gameMoney,strItemInfo)
		Dim i,j,k
		Dim sphn,blsResult : blsResult = False
		Dim rows,row,cols,col
		Dim itemHash,stackedCount,instance,itemExpire
		Dim cashStampHash,status,mailId
		cashStampHash = 0
		status = 8
		mailId = 0
		
		Dim conn
		set conn = Server.CreateObject("ADODB.Connection")	
		conn.Open(characterDBconnectionString)
	
		Dim itemSerial
		itemSerial = conn.Execute("select dbo.GetItemSerial(getdate())")(0)
		
		conn.Close
		
		Set sphn = new SPHelper_NoTran
		with sphn
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = characterDBconnectionString
			.SPName = "dbo.SendMail"
			Call .InitCommand()
			Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
			Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
			Call .AppendParam("@sender",adVarWChar,adParamInput,50,sender)
			Call .AppendParam("@title",adVarWChar,adParamInput,20,title)
			Call .AppendParam("@message",adVarWChar,adParamInput,150,message)
			Call .AppendParam("@gameMoney",adBigInt,adParamInput,,gameMoney)
			Call .AppendParam("@cashStampHash",adInteger,adParamInput,,cashStampHash)
			Call .AppendParam("@state",adUnsignedTinyInt,adParamInput,,status)
			Call .AppendParam("@mailId",adBigInt,adParamOutput,,null)
			blsResult = .ExecNoRecords()
		End with
		
		If blsResult Then
			Dim ret
			ret = sphn.GetParamValue("RETURN_VALUE")
			if ret<>0 Then
				SendMail = 0
				Exit Function
			End If
			mailId = sphn.GetParamValue("@mailId")
		End If
		set sphn = Nothing
		
		rows = Split(strItemInfo,"|")	
		If IsArray(rows) Then
			For i=0 To Ubound(rows)
				row = rows(i)
				cols = Split(row,",")			
				itemHash = cols(0)
				stackedCount = cols(1)
				instance = cols(2)
				itemExpire = cols(3)
				itemSerial = CStr(itemSerial) + 1
				
				Set sphn = new SPHelper_NoTran
				with sphn
					.DEBUG = False
					Set .cmd = Command
					.ConnStr = characterDBconnectionString
					.SPName = "dbo.SendMailItem"
					Call .InitCommand()
					Call .AppendParam("@mailId",adBigInt,adParamInput,,mailId)
					Call .AppendParam("@itemHash",adInteger,adParamInput,,itemHash)
					Call .AppendParam("@stackedCount",adSmallInt,adParamInput,,stackedCount)
					Call .AppendParam("@itemSerial",adBigInt,adParamInput,,itemSerial)
					Call .AppendParam("@instance",adBigInt,adParamInput,,instance)
					Call .AppendParam("@itemExpire",adDBTimeStamp,adParamInput,,FormatDt(itemExpire,"SQL_TM"))
					blsResult = .ExecNoRecords()
				End with
				set sphn = Nothing
			Next
		End If
		
		SendMail = mailId
	End Function
%>