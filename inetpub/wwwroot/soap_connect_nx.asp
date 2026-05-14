<%
	Class SoapConnect_NX
	'-------------------------------------
	' Member DEFINE
	'-------------------------------------
	Private blsDebug ' Debug모드	
	Private strSoapWSDL
	Private strSoapAction
	Private strSoapMethod
	Private strSoapNS
	Private strErrorMessage

	Private objConnector
	Private objSerializer
	Public objReader	'외부에서 접근
	Private objXML
	Private nodeFrg
	
	Private node_H
	Private node_SH
	
	'-------------------------------------
	' Public Property DEFINE
	'-------------------------------------
	Public Property Let Debug(value)
		blsDebug = value
	End Property

	Public Property Let WSDL(value)
		strSoapWSDL = value
	End Property
	
	Public Property Let Action(value)
		strSoapAction = value
	End Property
	
	Public Property Let Method(value)
		strSoapMethod = value
	End Property
	
	Public Property Let NS(value)
		strSoapNS = value
	End Property
	
	Public Property Get ErrorMessage()
		ErrorMessage = strErrorMessage
	End Property
	
	'-------------------------------------
	' Object Init
	'-------------------------------------
	Private Sub Class_Initialize()
		
		Set objConnector = Server.CreateObject("MSSOAP.HttpConnector30")
		Set objSerializer = Server.CreateObject("MSSOAP.SoapSerializer30")
		Set objReader = Server.CreateObject("MSSOAP.SoapReader30")
		
		Set objXML = Server.CreateObject("Msxml2.DOMDocument.3.0")
		objXML.async = False
		objXML.validateOnparse = False
		
		Set nodeFrg = objXML.createDocumentFragment()

	End Sub
	
	'-------------------------------------
	' Object Terminate
	'-------------------------------------
	Private Sub Class_Terminate()

		Set objConnector = Nothing
		Set objSerializer = Nothing
		Set objReader = Nothing
		Set objXML = Nothing
		
	End Sub
	
	'-------------------------------------
	' Private Method - Begin
	'-------------------------------------
	Private Sub MakeSoapXml()
		
		Dim node
		Set node = objXML.createElement(strSoapMethod)
		node.setAttribute "xmlns",strSoapNS

		If Not nodeFrg Is Nothing Then
			node.appendChild nodeFrg
		End If
		
		objXML.appendChild node

	End Sub
	'-------------------------------------
	' Private Method - End
	'-------------------------------------

	'-------------------------------------
	' Public Method - Begin
	'-------------------------------------
	Public Sub AppendSoapXmlByValue(name,value)

		On Error Resume Next

		Dim node
		Set node = objXML.createElement(name)
		node.Text = value
		nodeFrg.appendChild node

		If Err.Number<>0 Then
			strErrorMessage = Err.Description
			If blsDebug Then Response.write "ErrorMessage : " & strErrorMessage &"<br>"& vbCrLf
			Exit Sub
		End If

	End Sub
	
	Public Sub AppendSoapXmlByXml(name,xmlCh)

		On Error Resume Next

		Dim node,nodeCh,objXMLCh,i
		Set objXMLCh = Server.CreateObject("Msxml2.DOMDocument.3.0")
		objXMLCh.async = False
		objXMLCh.validateOnparse = False
		objXMLCh.loadXML("<child>"& xmlCh &"</child>")		
		Set node = objXML.createElement(name)	
		For i=0 To objXMLCh.documentElement.childNodes.length-1
			Set nodeCh = objXML.createElement(objXMLCh.documentElement.childNodes(i).nodeName)
			nodeCh.Text = objXMLCh.documentElement.childNodes(i).Text
			node.appendChild nodeCh
		Next
		Set objXMLCh = Nothing
		nodeFrg.appendChild node
		
		If Err.Number<>0 Then
			strErrorMessage = Err.Description
			If blsDebug Then Response.write "ErrorMessage : " & strErrorMessage &"<br>"& vbCrLf
			Exit Sub
		End If

	End Sub
	
	Public Sub InitializeSoap
		Dim oMD5
		Set oMD5 = new MD5
		
		Dim node

		Set node_H = objXML.createElement("Header") 
		node_H.setAttribute "xmlns",strSoapNS
		
		Set node = objXML.createElement("TS")
		node.Text = FormatDt(Now(),"ISO_TM")
		node_H.appendChild node
		
		Set node = objXML.createElement("IPAddress")
		node.Text = Request.ServerVariables("REMOTE_ADDR")
		node_H.appendChild node
		
		Set node = objXML.createElement("MD5Hash")
		node.Text = oMD5.MD5(Request.ServerVariables("REMOTE_ADDR") &"{82CFDDC1-E25F-41DD-8E2E-E93A4BF2C576}")
		node_H.appendChild node
		
		Set node_SH = objXML.createElement("SecurityHeader") 
		node_SH.setAttribute "xmlns",strSoapNS
		Set node = objXML.createElement("Created")
		node_SH.appendChild node
		Set node = objXML.createElement("HashKey")
		node_SH.appendChild node
		Set node = objXML.createElement("Data")
		node_SH.appendChild node
		Set node = objXML.createElement("Key")
		node.Text= 0
		node_SH.appendChild node
		Set node = objXML.createElement("CheckSum")
		node.Text= 0
		node_SH.appendChild node
		
		If blsDebug Then
			Response.write "---------------------------------------------"&"<br>"& vbCrLf
			Response.write "Debug InitializeSoap Params"&"<br>"& vbCrLf
			Response.write "---------------------------------------------"&"<br>"& vbCrLf
			Response.write "WSDL : " & strSoapWSDL &"<br>"& vbCrLf
			Response.write "Action : " & "https://nxsoap.nexon.com/remoting_version_2/InitializeSoap" &"<br>"& vbCrLf
			Response.write "Method : " & "InitializeSoap" &"<br>"& vbCrLf
			Response.write "NS : " & strSoapNS &"<br>"& vbCrLf
			Response.write "node_H : " & Replace(Replace(node_H.xml,"<","&lt;"),">","&gt;") &"<br>"& vbCrLf
			Response.write "node_SH : " & Replace(Replace(node_SH.xml,"<","&lt;"),">","&gt;") &"<br>"& vbCrLf
		End If
		
		objConnector.Property("EndPointURL") = strSoapWSDL
		objConnector.Property("SoapAction") = "https://nxsoap.nexon.com/remoting_version_2/InitializeSoap"
		objConnector.Connect
		objConnector.BeginMessage
		objSerializer.Init objConnector.InputStream
		objSerializer.StartEnvelope
		objSerializer.StartHeader
		objSerializer.WriteXML node_H.xml
		objSerializer.WriteXML node_SH.xml
		objSerializer.EndHeader
		objSerializer.StartBody
		objSerializer.EndBody
		objSerializer.EndEnvelope
		objConnector.EndMessage
		objReader.Load objConnector.OutputStream
		
		If blsDebug Then
			'Response.Write "RetXml : "& (objReader.DOM.xml) &"<br>"& vbCrLf
			Response.write "RetXml : " & Replace(Replace(objReader.DOM.xml,"<","&lt;"),">","&gt;") &"<br>"& vbCrLf
		End If
		
		Dim xmlDom
		Set xmlDom = Server.CreateObject("Msxml2.DOMDocument.3.0")
		xmlDom.load(objReader.DOM)
		Set node_SH = xmlDom.selectSingleNode("//SecurityHeader")
		node_SH.removeAttribute "soap:mustUnderstand"
		
		If blsDebug Then
			'Response.Write "RetXml : "& (objReader.DOM.xml) &"<br>"& vbCrLf
			Response.write "node_SH : " & Replace(Replace(node_SH.xml,"<","&lt;"),">","&gt;") &"<br>"& vbCrLf
		End If
	End Sub

	Public Sub Connect()
		
		'On Error Resume Next
		
		Call MakeSoapXml()
		
		If blsDebug Then
			Response.write "---------------------------------------------"&"<br>"& vbCrLf
			Response.write "Debug Connect Params"&"<br>"& vbCrLf
			Response.write "---------------------------------------------"&"<br>"& vbCrLf
			Response.write "WSDL : " & strSoapWSDL &"<br>"& vbCrLf
			Response.write "Action : " & strSoapAction &"<br>"& vbCrLf
			Response.write "Method : " & strSoapMethod &"<br>"& vbCrLf
			Response.write "NS : " & strSoapNS &"<br>"& vbCrLf
			Response.write "node_SH : " & Replace(Replace(node_SH.xml,"<","&lt;"),">","&gt;") &"<br>"& vbCrLf
			Response.write "Xml : " & Replace(Replace(objXML.xml,"<","&lt;"),">","&gt;") &"<br>"& vbCrLf
		End If

		objConnector.Property("EndPointURL") = strSoapWSDL
		objConnector.Property("SoapAction") = strSoapAction
		objConnector.Connect
		objConnector.BeginMessage
		objSerializer.Init objConnector.InputStream
		objSerializer.StartEnvelope
		objSerializer.StartHeader
		objSerializer.WriteXML node_SH.xml
		objSerializer.EndHeader
		objSerializer.StartBody
		objSerializer.WriteXML objXML.xml
		objSerializer.EndBody
		objSerializer.EndEnvelope
		objConnector.EndMessage
		objReader.Load objConnector.OutputStream
		
		If blsDebug Then
			'Response.Write "RetXml : "& (objReader.DOM.xml) &"<br>"& vbCrLf
			Response.write "RetXml : " & Replace(Replace(objReader.DOM.xml,"<","&lt;"),">","&gt;") &"<br>"& vbCrLf
		End If
		
		If Err.Number<>0 Then
			strErrorMessage = Err.Description
			If blsDebug Then Response.write "ErrorMessage : " & strErrorMessage &"<br>"& vbCrLf
			Exit Sub
		End If
		
	End Sub
	
	Public Function GetSoapXml()

		On Error Resume Next
		
		GetSoapXml = objXML.xml

		If Err.Number<>0 Then
			strErrorMessage = Err.Description
			GetSoapXml = "False"
			If blsDebug Then Response.write "ErrorMessage : " & strErrorMessage &"<br>"& vbCrLf
			Exit Function
		End If
    
	End Function
	
	Public Function GetArrayFromXml(nodeSchema,nodeDataset)

		Dim i,j
		Dim node
		ReDim tempArray(nodeDataset.length-1,nodeSchema.childNodes.length-1)
		For i=0 To nodeDataset.length-1
			For j=0 To nodeSchema.childNodes.length-1
				Set node = nodeDataset.Item(i).selectSingleNode(nodeSchema.childNodes(j).getAttribute("name"))
				If Not node Is Nothing Then
					tempArray(i,j) = node.Text
				End If
				Set node = Nothing
			Next
		Next		
		GetArrayFromXml = tempArray

	End Function
	'-------------------------------------
	' Public Method - End
	'-------------------------------------
	
	End Class
%>