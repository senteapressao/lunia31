<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	' request	: index,{characterName,hash,stackedCount,itemSerial,price(total),NeedItemId,NeedItemCount}
	' response	: 
	
	' preparing request parameters
	Dim rowData,index
	rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)

	If UBound(Parameters) < 1 Then Call Error("not enough parameter")	
	index = Parameters(0)
	
	If index = 0 Then
		Call Ok(0)
		response.end()
	End If
	
	dim arr,sql,i,count
	For count = 1 To index Step 1
		arr = Split(Parameters(count),SUBSEPARATOR)
	
		If ((UBound(arr)+1) Mod 7)<>0 Then
			Call Error("invalid number of item parameter")
		End If
		
		sql = sql &"exec BarterBuyLog"&_
			" "& SqlQuot(arr(0)) &_
			","& arr(1) &_
			","& arr(2) &_
			","& arr(3) &_
			","& arr(4) &_
			","& arr(5) &_
			","& arr(6) & ";"

	Next

	
	dim conn,rs
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open(characterDBconnectionString)
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.Open sql,conn,adOpenStatic,adLockReadOnly,adCmdText
	For i=1 To index
		set rs = rs.NextRecordset
	Next
	conn.Close()

	Call OK(0)
%>