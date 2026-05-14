<%
	Class SPHelper
	'-------------------------------------
	' Member DEFINE
	'-------------------------------------
	Public DEBUG

	Public cmd
	Public rs
	Public rs2

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

		ExecRecordset = True
	End Function

	Public Function ExecRecordset2()
		If DEBUG Then Call PrintCmdDebug()

		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.CursorLocation = adUseClient		
		rs.Open Command,,adOpenStatic,adLockReadOnly,adCmdStoredProc
		Set rs2 = rs.NextRecordset 
		Set rs.ActiveConnection = Nothing

		frk_n4ErrorCode = GetParamValue("@frk_n4ErrorCode")
		frk_strErrorText = GetParamValue("@frk_strErrorText")

		If DEBUG Then Response.Write frk_strErrorText &"<br />"
		
		If frk_n4ErrorCode<>0 Then
			ExecRecordset2 = False
			Exit Function
		End If

		ExecRecordset2 = True
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

	Public Sub PrintRS(RS)
		Dim i,j,k
		If TypeName(RS)="Recordset" Then
			Do Until RS.Eof
				If i=0 Then
					Response.Write "<table border='1'>"
					Response.Write "<tr>"
					For Each j In RS.Fields
						Response.Write "<td style='font-size:11px;'>"
						Response.Write j.Name
						Response.Write "</td>"
					Next
					Response.Write "</tr>"
				End If
				Response.Write "<tr height='16'>"
				For Each j In RS.Fields
					Response.Write "<td style='font-size:11px;'>"
					Response.Write RS(j.Name)
					Response.Write "</td>"
				Next
				Response.Write "</tr>"
				RS.MoveNext : i=i+1
				If RS.Eof Then Response.Write "</table>"
			Loop
		End If
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