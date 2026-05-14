<%

Public Function URLDecode(sEncodedURL)
	Dim iLoop
	Dim sRtn
	Dim sTmp

	If Len(sEncodedURL) > 0 Then
		' Loop through each char


		For iLoop = 1 To Len(sEncodedURL)
			sTmp = Mid(sEncodedURL, iLoop, 1)
			sTmp = Replace(sTmp, "+", " ")
			' If char is % then get next two chars
			' and convert from HEX to decimal


			If sTmp = "%" and LEN(sEncodedURL) + 1 > iLoop + 2 Then
				sTmp = Mid(sEncodedURL, iLoop + 1, 2)
				sTmp = Chr(CInt("&H" & sTmp))
				' Increment loop by 2
				iLoop = iLoop + 2
			End If
			sRtn = sRtn & sTmp
		Next
		URLDecode = sRtn
	End If
End Function

%>