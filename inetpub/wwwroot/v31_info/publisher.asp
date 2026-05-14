<!--#include file="./soap_connect.asp"-->
<%
	Class Publisher
	'-------------------------------------
	' Member DEFINE
	'-------------------------------------
	Public DEBUG
	Public strErrorText
	
	Public strReturn

	Private soap
	Private xmlDom
	Private strSoapWSDL
	Private	strSoapAction
	Private	strSoapNS
	Private	strSoapMethod

	'-------------------------------------
	' Public Property DEFINE
	'-------------------------------------

	'-------------------------------------
	' Object Init
	'-------------------------------------
	Private Sub Class_Initialize()
		DEBUG = False

		Set soap = new SoapConnect
		Set xmlDom = Server.CreateObject("Msxml2.DOMDocument.3.0")

	End Sub

	'-------------------------------------
	' Object Terminate
	'-------------------------------------
	Private Sub Class_Terminate()

		Set soap = Nothing
		Set xmlDom = Nothing

	End Sub

	'-------------------------------------
	' Public Method - Begin
	'-------------------------------------

	Public Function get_money(strUserID)
		On Error Resume Next

		If DEBUG Then
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "Debug get_money Params"&"<br />"& vbCrLf
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "id : "& strUserID &"<br />"& vbCrLf
		End If
		
		Dim nodeResult
		Dim nErrorCode,strErrorMessage

		strSoapWSDL = LoginWSDL
		strSoapAction = "http://tempuri.org/check_id_point"
		strSoapMethod = "check_id_point"
		strSoapNS = "http://tempuri.org/"
		
		With soap
			.Debug = DEBUG
			.WSDL = strSoapWSDL
			.Action = strSoapAction
			.Method = strSoapMethod
			.NS = strSoapNS
			Call .AppendSoapXmlByValue("gameid",strUserID)
			.Connect
		End With

		If Not(soap.objReader.Fault Is Nothing) Then
			If DEBUG Then Response.write "objReader.FaultString : "& soap.objReader.FaultString.Text &"<br />"& vbCrLf

			strErrorText = "error occurs during soap connect"
			get_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If

		Set xmlDom = CreateObject("Msxml2.DOMDocument.3.0")
		xmlDom.async = False
		xmlDom.validateOnparse = False
		xmlDom.load(soap.objReader.DOM)
		
		Set nodeResult = xmlDom.selectSingleNode("//"& strSoapMethod &"Result")
		If nodeResult Is Nothing Then
			strErrorText = "return value is not valid"
			get_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If
		
		strReturn = nodeResult.Text
		
		strErrorText = "success"
		get_money = True

	End Function

	Public Function use_money(strUserID,nTotalPrice,TID,oidProduct,nOrderQuantity,strProductName,nSalePrice _
							,strType,strUserID_R)
		'On Error Resume Next
		
		If DEBUG Then
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "Debug use_money Params"&"<br />"& vbCrLf
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "gameid : "& strUserID &"<br />"& vbCrLf
			Response.write "calpoint : "& nTotalPrice &"<br />"& vbCrLf
			Response.write "itemno : "& oidProduct &"<br />"& vbCrLf
			Response.write "itemsize : "& nOrderQuantity &"<br />"& vbCrLf
		End If
		
		Dim nodeResult
		Dim nErrorCode,strErrorMessage	

		strSoapWSDL = LoginWSDL		
		strSoapAction = "http://tempuri.org/buy_lunia"
		strSoapMethod = "buy_lunia" 
		strSoapNS = "http://tempuri.org/"
		
		With soap
			.Debug = DEBUG
			.WSDL = strSoapWSDL
			.Action = strSoapAction
			.Method = strSoapMethod
			.NS = strSoapNS
			Call .AppendSoapXmlByValue("gameid",strUserID)
			Call .AppendSoapXmlByValue("calpoint",nTotalPrice)
			Call .AppendSoapXmlByValue("itemno",oidProduct)
			Call .AppendSoapXmlByValue("itemsize",nOrderQuantity)
			.Connect
		End With

		If Not(soap.objReader.Fault Is Nothing) Then
			If DEBUG Then Response.write "objReader.FaultString : "& soap.objReader.FaultString.Text &"<br />"& vbCrLf

			strErrorText = "error occurs during soap connect"
			use_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If


		Set xmlDom = CreateObject("Msxml2.DOMDocument.3.0")
		xmlDom.async = False
		xmlDom.validateOnparse = False
		xmlDom.load(soap.objReader.DOM)
		
		Set nodeResult = xmlDom.selectSingleNode("//"& strSoapMethod &"Result")
		If nodeResult Is Nothing Then
			strErrorText = "return value is not valid"
			use_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If

		If Left(nodeResult.Text,1)="0" Then
			Dim arrResult
			arrResult = Split(nodeResult.Text,"|")
			If IsArray(arrResult)=False Then
				strErrorText = "return value is not valid"
				use_money = False
				If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
				Exit Function
			End If
			If Ubound(arrResult)<3 Then
				strErrorText = "return value is not valid"
				use_money = False
				If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
				Exit Function
			End If
			strReturn = arrResult(3)
		Else
			strErrorText = "return value is not valid"
			use_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If
		
		strErrorText = "success"
		use_money = True

	End Function

	Public Function rollback_money(strUserID,oidPayment,TID,nTotalPrice)
		'On Error Resume Next

		If DEBUG Then
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "Debug rollback_money Params"&"<br />"& vbCrLf
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "gameid : "& strUserID &"<br />"& vbCrLf
			Response.write "addpoint : "& nTotalPrice &"<br />"& vbCrLf
			Response.write "unicode : "& oidPayment &"<br />"& vbCrLf
		End If
		
		Dim nodeResult
		Dim nErrorCode,strErrorMessage

		strSoapWSDL = LoginWSDL
		strSoapAction = "http://tempuri.org/return_lunia"
		strSoapMethod = "return_lunia"
		strSoapNS = "http://tempuri.org/"
		
		With soap
			.Debug = DEBUG
			.WSDL = strSoapWSDL
			.Action = strSoapAction
			.Method = strSoapMethod
			.NS = strSoapNS
			Call .AppendSoapXmlByValue("gameid",strUserID)
			Call .AppendSoapXmlByValue("addpoint",nTotalPrice)
			Call .AppendSoapXmlByValue("unicodeno",oidPayment)
			.Connect
		End With

		If Not(soap.objReader.Fault Is Nothing) Then
			If DEBUG Then Response.write "objReader.FaultString : "& soap.objReader.FaultString.Text &"<br />"& vbCrLf

			strErrorText = "error occurs during soap connect"
			rollback_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If
		
		Set xmlDom = CreateObject("Msxml2.DOMDocument.3.0")
		xmlDom.async = False
		xmlDom.validateOnparse = False
		xmlDom.load(soap.objReader.DOM)
		
		Set nodeResult = xmlDom.selectSingleNode("//"& strSoapMethod &"Result")
		If nodeResult Is Nothing Then
			strErrorText = "return value is not valid"
			rollback_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If
		
		If Left(nodeResult.Text,1)="0" Then
			strReturn = 0
		Else
			strErrorText = "return value is not valid"
			rollback_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If

		strReturn = nErrorCode
		
		strErrorText = "success"
		rollback_money = True

	End Function
	
	'-------------------------------------
	' Public Method - End
	'-------------------------------------
		
	End Class
%>