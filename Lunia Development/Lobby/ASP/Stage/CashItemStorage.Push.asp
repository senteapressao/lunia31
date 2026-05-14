<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName,itemSerial,itemHash,stackedCount,instance,pageIndex,position
	' response : pageIndex,position

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False	
	Dim retString : retString = ""
	
	Dim characterName,itemSerial,itemHash,stackedCount,instance,pageIndex,position
	If UBound(parameters)<6 Then Call Error("not enough parameter")
	charactername = Parameters(0)
	itemSerial = Parameters(1)
	itemHash = Parameters(2)
	stackedCount = Parameters(3)
	instance = Parameters(4)
	pageIndex = Parameters(5)
	position = Parameters(6)
	
	Dim sql,affected
	sql = "if not exists(select 'x' from dbo.CashItemStorage (nolock) where characterName=? and pageIndex=? and position=?)"&_
		" insert dbo.CashItemStorage(characterName,pageIndex,position,itemSerial,itemHash,stackedCount,instance)"&_
		" select ?,?,?,?,?,?,?"

	with Command
		.ActiveConnection = characterDBconnectionString
		.CommandText = sql
		.CommandType = adCmdText
		
		.Parameters.Append .CreateParameter("@characterName",adVarWChar,adParamInput,50,characterName)
		.Parameters.Append .CreateParameter("@pageIndex",adUnsignedTinyInt,adParamInput,,pageIndex)
		.Parameters.Append .CreateParameter("@position",adUnsignedTinyInt,adParamInput,,position)
		
		.Parameters.Append .CreateParameter("@characterName",adVarWChar,adParamInput,50,characterName)
		.Parameters.Append .CreateParameter("@pageIndex",adUnsignedTinyInt,adParamInput,,pageIndex)
		.Parameters.Append .CreateParameter("@position",adUnsignedTinyInt,adParamInput,,position)
		.Parameters.Append .CreateParameter("@itemSerial",adBigInt,adParamInput,,itemSerial)
		.Parameters.Append .CreateParameter("@itemHash",adInteger,adParamInput,,itemHash)
		.Parameters.Append .CreateParameter("@stackedCount",adSmallInt,adParamInput,,stackedCount)
		.Parameters.Append .CreateParameter("@instance",adBigInt,adParamInput,,instance)		
		
		.Execute affected
	End with
	
	If affected<>1 Then
		Call Error(1)
	End If
	
	retString = retString & pageIndex & SEPARATOR & position

	Call Ok(retString)
%>
