<%
	Dim SEPARATOR,SUBSEPARATOR
	Dim Parameters
	SUBSEPARATOR=Chr(11)
	SEPARATOR=Chr(8)

	Parameters = Split(Request.QueryString,SEPARATOR)

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

%>

