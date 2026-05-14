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

	Public isNetcafe

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

		strSoapWSDL = LoginWSDL

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
	Public Function auth_user(strUserID,strPassword,strUserIP)
		On Error Resume Next
		
		If DEBUG Then
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "Debug auth_user Params"&"<br />"& vbCrLf
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "WSDL : "& strSoapWSDL &"<br />"& vbCrLf
			Response.write "accountName : "& strUserID &"<br />"& vbCrLf
			Response.write "password : "& strPassword &"<br />"& vbCrLf
			Response.write "clientIP : "& strUserIP &"<br />"& vbCrLf
		End If
		
		Dim nodeResult
		Dim nErrorCode,strErrorMessage
	
		strSoapAction = "http://playmojo.com/auth_member"
		strSoapMethod = "auth_member"
		strSoapNS = "http://playmojo.com/"
		
		With soap
			.Debug = DEBUG
			.WSDL = strSoapWSDL
			.Action = strSoapAction
			.Method = strSoapMethod
			.NS = strSoapNS
			Call .AppendSoapXmlByValue("accountName",strUserID)
			Call .AppendSoapXmlByValue("password",strPassword)
			Call .AppendSoapXmlByValue("clientIP",strUserIP)
			.Connect
		End With

		If Not(soap.objReader.Fault Is Nothing) Then
			If DEBUG Then Response.write "objReader.FaultString : "& soap.objReader.FaultString.Text &"<br />"& vbCrLf

			strErrorText = "error occurs during soap connect"
			auth_user = False
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
			auth_user = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If	

		nErrorCode = nodeResult.Text

		If nErrorCode<>0 Then
			strErrorText = "ErrorCode : "& nErrorCode &" / ErrorMsg : "& strErrorMessage
			strReturn = nErrorCode
			auth_user = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If

		isNetCafe = xmlDom.selectSingleNode("//isNetCafe").Text

		If Err.number<>0 Then
			strErrorText = "ErrorCode : "& Err.number &" / ErrorMsg : "& Err.Description
			strReturn = nErrorCode
			auth_user = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If

		On Error Goto 0		

		strReturn = nErrorCode

		strErrorText = "success"
		auth_user = True

	End Function

	Public Function get_money(strUserID)
		If BillingMode="None" Then
			strReturn = 1000
			strErrorText = "success"
			get_money = True
			Exit Function
		End If

		On Error Resume Next

		If DEBUG Then
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "Debug get_money Params"&"<br />"& vbCrLf
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "strGUserID : "& strUserID &"<br />"& vbCrLf
		End If
		
		Dim nodeResult
		Dim nErrorCode,strErrorMessage

		strSoapWSDL = ShopWSDL
		strSoapAction = ShopNS &"Lunia/LuniaWebService/GetUserBalance"
		strSoapMethod = "GetUserBalance" 
		strSoapNS = ShopNS &"Lunia/LuniaWebService/"

		With soap
			.Debug = DEBUG
			.WSDL = strSoapWSDL
			.Action = strSoapAction
			.Method = strSoapMethod
			.NS = strSoapNS
			Call .AppendSoapXmlByValue("strGameCode","LN")
			Call .AppendSoapXmlByValue("strGUserID",strUserID)
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
		
		nErrorCode = xmlDom.selectSingleNode("//intRetVal").Text
		strErrorMessage = xmlDom.selectSingleNode("//strErrMsg").Text

		If nErrorCode<>0 Then
			strErrorText = "ErrorCode : "& nErrorCode &" / ErrorMsg : "& strErrorMessage
			get_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If

		Dim intGMReal,intGMBonus

		intGMReal = xmlDom.selectSingleNode("//intGMReal").Text
		intGMBonus = xmlDom.selectSingleNode("//intGMBonus").Text

		If Err.number<>0 Then
			strErrorText = "ErrorCode : "& Err.number &" / ErrorMsg : "& Err.Description
			strReturn = nErrorCode
			get_money= False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If

		On Error Goto 0

		strReturn = CLng(intGMReal) + CLng(intGMBonus)
		
		strErrorText = "success"
		get_money = True

	End Function

	Public Function use_money(strUserID,nTotalPrice,TID,oidProduct,nOrderQuantity,strProductName,nSalePrice _
							,strType,strUserID_R)
		If BillingMode="None" Then
			strErrorText = "success"
			use_money = True
			Exit Function
		End If
		
		'On Error Resume Next

		Dim nPurchaseType
		If strType = "buy" Then
			nPurchaseType = 1
		ElseIf strType = "present" Then
			nPurchaseType = 2
		End If
		
		If DEBUG Then
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "Debug use_money Params"&"<br />"& vbCrLf
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "strGameCode : LN<br />"& vbCrLf
			Response.write "strGUserID : "& strUserID &"<br />"& vbCrLf
			Response.write "intPurchaseType : "& nPurchaseType &"<br />"& vbCrLf
			Response.write "strRGUserID : "& strUserID_R &"<br />"& vbCrLf
			Response.write "strGameItemID : "& oidProduct &"<br />"& vbCrLf
			Response.write "strItemName : "& strProductName &"<br />"& vbCrLf
			Response.write "intItemCnt : "& nOrderQuantity &"<br />"& vbCrLf
			Response.write "intItemPrice : "& nSalePrice &"<br />"& vbCrLf
			Response.write "intChargeAmt : "& nTotalPrice &"<br />"& vbCrLf
			Response.write "strGameOrderNo : "& TID &"<br />"& vbCrLf
		End If
		
		Dim nodeResult
		Dim nErrorCode,strErrorMessage	

		strSoapWSDL = ShopWSDL
		strSoapAction = ShopNS &"Lunia/LuniaWebService/GameItemCharge"
		strSoapMethod = "GameItemCharge"
		strSoapNS = ShopNS &"Lunia/LuniaWebService/"
		
		With soap
			.Debug = DEBUG
			.WSDL = strSoapWSDL
			.Action = strSoapAction
			.Method = strSoapMethod
			.NS = strSoapNS
			Call .AppendSoapXmlByValue("strGameCode","LN")
			Call .AppendSoapXmlByValue("strGUserID",strUserID)
			Call .AppendSoapXmlByValue("intPurchaseType",nPurchaseType)	'1:Buy himself, 2:Give a Present
			Call .AppendSoapXmlByValue("strRGUserID",strUserID_R)
			Call .AppendSoapXmlByValue("strGameItemID",oidProduct)
			Call .AppendSoapXmlByValue("strItemName",strProductName)
			Call .AppendSoapXmlByValue("intItemCnt",nOrderQuantity)
			Call .AppendSoapXmlByValue("intItemPrice",nSalePrice)
			Call .AppendSoapXmlByValue("intChargeAmt",nTotalPrice)
			Call .AppendSoapXmlByValue("strGameOrderNo",TID)
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

		nErrorCode = xmlDom.selectSingleNode("//intRetVal").Text
		strErrorMessage = xmlDom.selectSingleNode("//strErrMsg").Text
		
		If nErrorCode<>0 Then
			strErrorText = "ErrorCode : "& nErrorCode &" / ErrorMsg : "& strErrorMessage
			use_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If

		Dim intPurchaseNo,intChargedAmt,intGMReal,intGMBonus

		intPurchaseNo = xmlDom.selectSingleNode("//intPurchaseNo").Text
		intChargedAmt = xmlDom.selectSingleNode("//intChargedAmt").Text
		intGMReal = xmlDom.selectSingleNode("//intGMReal").Text
		intGMBonus = xmlDom.selectSingleNode("//intGMBonus").Text

		If Err.number<>0 Then
			strErrorText = "ErrorCode : "& Err.number &" / ErrorMsg : "& Err.Description
			strReturn = nErrorCode
			use_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If

		On Error Goto 0

		strReturn = intPurchaseNo
		
		strErrorText = "success"
		use_money = True

	End Function

	Public Function rollback_money(strUserID,oidPayment,TID,nTotalPrice)
		If BillingMode="None" Then
			strErrorText = "success"
			rollback_money = True
			Exit Function
		End If
		
		'On Error Resume Next

		If DEBUG Then
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "Debug rollback_money Params"&"<br />"& vbCrLf
			Response.write "---------------------------------------------"&"<br />"& vbCrLf
			Response.write "strGameCode : LN<br />"& vbCrLf
			Response.write "strGUserID : "& strUserID &"<br />"& vbCrLf
			Response.write "intPurchaseNo : "& oidPayment &"<br />"& vbCrLf
		End If
		
		Dim nodeResult
		Dim nErrorCode,strErrorMessage

		strSoapWSDL = ShopWSDL
		strSoapAction = ShopNS &"Lunia/LuniaWebService/GameItemChargeCancel"
		strSoapMethod = "GameItemChargeCancel"
		strSoapNS = ShopNS &"Lunia/LuniaWebService/"
		
		With soap
			.Debug = DEBUG
			.WSDL = strSoapWSDL
			.Action = strSoapAction
			.Method = strSoapMethod
			.NS = strSoapNS
			Call .AppendSoapXmlByValue("strGameCode","LN")
			Call .AppendSoapXmlByValue("strGUserID",strUserID)
			Call .AppendSoapXmlByValue("intPurchaseNo",oidPayment)
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
		
		nErrorCode = xmlDom.selectSingleNode("//intRetVal").Text
		strErrorMessage = xmlDom.selectSingleNode("//strErrMsg").Text
		
		If nErrorCode<>0 Then
			strErrorText = "ErrorCode : "& nErrorCode &" / ErrorMsg : "& strErrorMessage
			rollback_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If

		Dim intGMReal,intGMBonus,intCancelledAmt,strGameOrderNo

		intGMReal = xmlDom.selectSingleNode("//intGMReal").Text
		intGMBonus = xmlDom.selectSingleNode("//intGMBonus").Text
		intCancelledAmt = xmlDom.selectSingleNode("//intCancelledAmt").Text
		strGameOrderNo = xmlDom.selectSingleNode("//strGameOrderNo").Text

		If Err.number<>0 Then
			strErrorText = "ErrorCode : "& Err.number &" / ErrorMsg : "& Err.Description
			strReturn = nErrorCode
			rollback_money = False
			If DEBUG Then Response.write "Error : "& strErrorText &"<br />"& vbCrLf
			Exit Function
		End If

		On Error Goto 0

		strReturn = strGameOrderNo
		
		strErrorText = "success"
		rollback_money = True

	End Function
	
	'-------------------------------------
	' Public Method - End
	'-------------------------------------
		
	End Class
%>