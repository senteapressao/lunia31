<%
	Class SPHelper_NoTran
	'-------------------------------------
	' Member DEFINE
	'-------------------------------------
	Public DEBUG

	Public cmd
	Public rs
	Public rs2
	Public rs3
	Public rs4

	Public ConnStr
	Public SPName

	'-------------------------------------
	' Object Init
	'-------------------------------------
	Private Sub Class_Initialize()
		DEBUG = False
	End Sub

	'-------------------------------------
	' Object Terminate
	'-------------------------------------
	Private Sub Class_Terminate()
		TerminateCommand()
		Set rs = Nothing
		Set rs2 = Nothing
		Set rs3 = Nothing
		Set rs4 = Nothing
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

	End Sub

	Public Function ExecNoRecords()
		If DEBUG Then Call PrintCmdDebug()
				
		cmd.Execute ,,adExecuteNoRecords
		
		ExecNoRecords = True
	End Function

	Public Function ExecRecordset()
		If DEBUG Then Call PrintCmdDebug()
		
		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.CursorLocation = adUseClient
		rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdStoredProc

		If rs.State = adStateClosed Then
			ExecRecordset = False
			Exit Function
		End If
		
		ExecRecordset = True
	End Function
	
	Public Function ExecRecordset2()
		If DEBUG Then Call PrintCmdDebug()

		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.CursorLocation = adUseClient		
		rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdStoredProc
		
		If rs.State = adStateClosed Then
			ExecRecordset2 = False
			Exit Function
		End If
		
		Set rs2 = rs.NextRecordset
		Set rs.ActiveConnection = Nothing

		ExecRecordset2 = True
	End Function
	
	Public Function ExecRecordset3()
		If DEBUG Then Call PrintCmdDebug()

		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.CursorLocation = adUseClient		
		rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdStoredProc
		
		If rs.State = adStateClosed Then
			ExecRecordset3 = False
			Exit Function
		End If
		
		Set rs2 = rs.NextRecordset 
		Set rs3 = rs2.NextRecordset 
		Set rs.ActiveConnection = Nothing

		ExecRecordset3 = True
	End Function
	
	Public Function ExecRecordset4()
		If DEBUG Then Call PrintCmdDebug()

		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.CursorLocation = adUseClient		
		rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdStoredProc
		
		If rs.State = adStateClosed Then
			ExecRecordset4 = False
			Exit Function
		End If
		
		Set rs2 = rs.NextRecordset 
		Set rs3 = rs2.NextRecordset 
		Set rs4 = rs3.NextRecordset 
		Set rs.ActiveConnection = Nothing

		ExecRecordset4 = True
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