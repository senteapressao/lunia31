<%
	Function GetConnStr(SERVER,DBNAME,UID,PWD)
		GetConnStr = "DRIVER=SQL Server;SERVER="& SERVER &";DATABASE="& DBNAME &";Network=DBMSSOCN;Address="& SERVER &";User Id="& UID &";PASSWORD="& PWD &";"	
	End Function
	
	Sub DeleteParameters(cmd)
		Do Until cmd.Parameters.Count = 0
			cmd.Parameters.Delete (0)
		Loop
	End Sub
	
	Sub AppendParam(cmd,paramName,paramType,paramDirection,paramSize,paramValue)
		If IsEmpty(paramValue) OR paramValue="" Then
			paramValue = null
		End If
		cmd.Parameters.Append cmd.CreateParameter(paramName,paramType,paramDirection,paramSize,paramValue)
	End Sub
	
	Function GetParamValue(cmd,paramName)
		GetParamValue = cmd.Parameters(paramName).value
	End Function
	
	Sub TerminateCommand(cmd)
		Set cmd.ActiveConnection = Nothing
		Do Until cmd.Parameters.Count = 0
			cmd.Parameters.Delete (0)
		Loop
	End Sub
	
	Sub PrintADOConst()
		Response.Write "adStateClosed : "& adStateClosed &"<br />"
		
	End Sub

	Class SPHelper
	'-------------------------------------
	' Member DEFINE
	'-------------------------------------
	Public DEBUG

	Public cmd
	Public rs
	Public rs2
	Public rs3
	Public rs4
	Public rs5
	Public rs6
	Public rs7
	Public rs8
	Public rs9
	Public rs10
	Public rs11
	
	Public frk_n4ErrorCode
	Public frk_strErrorText
	Public frk_isRequiresNewTransaction
	Public ConnStr
	Public SPName

	'-------------------------------------
	' Object Init
	'-------------------------------------
	Private Sub Class_Initialize()
		DEBUG = False
		frk_isRequiresNewTransaction = 1
	End Sub

	'-------------------------------------
	' Object Terminate
	'-------------------------------------
	Private Sub Class_Terminate()
		TerminateCommand()
		Set rs = Nothing
		Set rs2 = Nothing
	End Sub
	
	'-------------------------------------
	' Public Method - Begin
	'-------------------------------------
	Public Sub InitCommand()
		with cmd
		.ActiveConnection = ConnStr
		.CommandText = SPName
		.CommandType = adCmdStoredProc
		End with

		Call AppendParam("@frk_n4ErrorCode",adInteger,adParamOutput,,null)
		Call AppendParam("@frk_strErrorText",adVarWChar,adParamOutput,100,null)
		Call AppendParam("@frk_isRequiresNewTransaction",adUnsignedTinyInt,adParamInput,,1)
	End Sub

	Public Function ExecNoRecords()
		If DEBUG Then Call PrintCmdDebug()
				
		cmd.Execute ,,adExecuteNoRecords

		frk_n4ErrorCode = GetParamValue("@frk_n4ErrorCode")
		frk_strErrorText = GetParamValue("@frk_strErrorText")

		If DEBUG Then Response.Write frk_strErrorText &"<br />"
		
		If frk_n4ErrorCode<>0 Then
			ExecNoRecords = False
			Exit Function
		End If

		ExecNoRecords = True
	End Function

	Public Function ExecRecordset()
		If DEBUG Then Call PrintCmdDebug()
		
		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.CursorLocation = adUseClient
		rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdStoredProc
		
		frk_n4ErrorCode = GetParamValue("@frk_n4ErrorCode")
		frk_strErrorText = GetParamValue("@frk_strErrorText")

		If DEBUG Then Response.Write frk_strErrorText &"<br />"
		
		If frk_n4ErrorCode<>0 Then
			ExecRecordset = False
			Exit Function
		End If
		
		Set rs.ActiveConnection = Nothing

		ExecRecordset = True
	End Function

	Public Function ExecRecordset2()
		If DEBUG Then Call PrintCmdDebug()

		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.CursorLocation = adUseClient		
		rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdStoredProc
		
		frk_n4ErrorCode = GetParamValue("@frk_n4ErrorCode")
		frk_strErrorText = GetParamValue("@frk_strErrorText")

		If DEBUG Then Response.Write frk_strErrorText &"<br />"
		
		If frk_n4ErrorCode<>0 Then
			ExecRecordset2 = False
			Exit Function
		End If
		
		Set rs2 = rs.NextRecordset 
		Set rs.ActiveConnection = Nothing

		ExecRecordset2 = True
	End Function
	
	Public Function ExecRecordset4()
		If DEBUG Then Call PrintCmdDebug()

		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.CursorLocation = adUseClient		
		rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdStoredProc
		
		frk_n4ErrorCode = GetParamValue("@frk_n4ErrorCode")
		frk_strErrorText = GetParamValue("@frk_strErrorText")

		If DEBUG Then Response.Write frk_strErrorText &"<br />"
		
		If frk_n4ErrorCode<>0 Then
			ExecRecordset4 = False
			Exit Function
		End If
		
		Set rs2 = rs.NextRecordset 
		Set rs3 = rs.NextRecordset 
		Set rs4 = rs.NextRecordset
		Set rs.ActiveConnection = Nothing

		ExecRecordset4 = True
	End Function
	
	Public Function ExecRecordset5()
		If DEBUG Then Call PrintCmdDebug()

		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.CursorLocation = adUseClient		
		rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdStoredProc
		
		frk_n4ErrorCode = GetParamValue("@frk_n4ErrorCode")
		frk_strErrorText = GetParamValue("@frk_strErrorText")

		If DEBUG Then Response.Write frk_strErrorText &"<br />"
		
		If frk_n4ErrorCode<>0 Then
			ExecRecordset5 = False
			Exit Function
		End If
		
		Set rs2 = rs.NextRecordset 
		Set rs3 = rs.NextRecordset 
		Set rs4 = rs.NextRecordset
		Set rs5 = rs.NextRecordset
		Set rs.ActiveConnection = Nothing

		ExecRecordset5 = True
	End Function
	
	Public Function ExecRecordset11()
		If DEBUG Then Call PrintCmdDebug()

		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.CursorLocation = adUseClient		
		rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdStoredProc
		
		frk_n4ErrorCode = GetParamValue("@frk_n4ErrorCode")
		frk_strErrorText = GetParamValue("@frk_strErrorText")

		If DEBUG Then Response.Write frk_strErrorText &"<br />"
		
		If frk_n4ErrorCode<>0 Then
			ExecRecordset11 = False
			Exit Function
		End If
		
		Set rs2 = rs.NextRecordset 
		Set rs3 = rs.NextRecordset 
		Set rs4 = rs.NextRecordset
		Set rs5 = rs.NextRecordset
		Set rs6 = rs.NextRecordset 
		Set rs7 = rs.NextRecordset 
		Set rs8 = rs.NextRecordset
		Set rs9 = rs.NextRecordset
		Set rs10 = rs.NextRecordset
		Set rs11 = rs.NextRecordset
		Set rs.ActiveConnection = Nothing

		ExecRecordset11 = True
	End Function
	
	Public Function ExecActiveRecordset()
		If DEBUG Then Call PrintCmdDebug()
		
		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.CursorLocation = adUseClient
		rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdStoredProc

		frk_n4ErrorCode = GetParamValue("@frk_n4ErrorCode")
		frk_strErrorText = GetParamValue("@frk_strErrorText")

		If DEBUG Then Response.Write frk_strErrorText &"<br />"
		
		If frk_n4ErrorCode<>0 Then
			ExecRecordset = False
			Exit Function
		End If

		ExecActiveRecordset = True
	End Function

	Public Sub AppendParam(paramName,paramType,paramDirection,paramSize,paramValue)
		If IsEmpty(paramValue) OR paramValue="" Then
			paramValue = null
		End If
		cmd.Parameters.Append cmd.CreateParameter(paramName,paramType,paramDirection,paramSize,paramValue)
	End Sub
	
	Public Function GetParamValue(paramName)
		GetParamValue = cmd.Parameters(paramName).value
	End Function
	
	Public Sub TerminateCommand()
		Set cmd.ActiveConnection = Nothing
		Do Until cmd.Parameters.Count = 0
			cmd.Parameters.Delete (0)
		Loop
	End Sub	

	'-------------------------------------
	' Private Method - Begin
	'-------------------------------------
	Private Sub PrintCmdDebug()
		Dim i
		Response.Write SPName &"<p />"		
		Response.Write "<table border='1'>"
		Response.Write "<tr>"
		Response.Write "<td>paramName</td>"
		Response.Write "<td>paramType</td>"
		Response.Write "<td>paramDirection</td>"
		Response.Write "<td>paramSize</td>"
		Response.Write "<td>paramValue</td>"
		Response.Write "</tr>"
		For i=0 To cmd.Parameters.Count-1
			Response.Write "<tr>"
			Response.Write "<td>"& cmd.Parameters(i).Name &"</td>"
			Response.Write "<td>"& cmd.Parameters(i).type &"</td>"
			Response.Write "<td>"& cmd.Parameters(i).Direction &"</td>"
			Response.Write "<td>"& cmd.Parameters(i).size &"</td>"
			Response.Write "<td>"& cmd.Parameters(i).value &"</td>"
			Response.Write "</tr>"
		Next
		Response.Write "</table>"
	End Sub

	End Class
%>