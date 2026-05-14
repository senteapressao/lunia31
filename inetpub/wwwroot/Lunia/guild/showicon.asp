<% @Language="VBScript" %>
<% Option Explicit %>
<%
	dim guildid,adoStream, FPath, fs,iconPath
	if Request.QueryString("guildid").Count = 0 then
		guildid = 0
	else
		guildid = Request.QueryString("guildid")
	end if
	
	Dim sResult : sResult = GetTextFromUrl("http://localhost/getGuildIcon/" & guildid)
	
	Response.ContentType = "image/png"
	
	Response.BinaryWrite sResult

	Response.End
	
	
	Function GetTextFromUrl(url)

		Dim oXMLHTTP
		Dim strStatusTest
	
		Set oXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
		
        if oXMLHTTP is nothing then 
			Set oXMLHTTP = CreateObject("Microsoft.XMLHTTP")
		end if
		oXMLHTTP.Open "GET", url, False
		oXMLHTTP.Send
	
		If oXMLHTTP.Status = 200 Then
			GetTextFromUrl =  oXMLHTTP.responseBody
		Else
			response.End
		End If

	End Function
%> 
