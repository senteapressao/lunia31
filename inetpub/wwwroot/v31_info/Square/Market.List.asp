<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : 
	' response : totalCount
	'			,{marketId,characterName,category,itemType,itemHash,stackedCount,itemSerial,instance,itemExpire
	'			,itemName,salePrice,expireDate,showLevel}
	'			,{marketId,characterName,category,itemType,itemHash,stackedCount,itemSerial,instance,itemExpire
	'			,itemName,petName,petLevel,isRare,rareProbability,enchantSerial,salePrice,expireDate,showLevel}

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim totalCount : totalCount = 0

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.Market_List"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		blsResult = .ExecRecordset()
	End with

	If blsResult Then
		Dim ret
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
		Set rs = sphn.rs
	Else
		ret = sphn.GetParamValue("RETURN_VALUE")
		if ret<>0 Then
			Call Error(ret)
		End If
	End If
	set sphn = Nothing
	
	If rs.Eof Then
		Response.Write "0"
		Response.End
	End If

	Dim arr
	arr = rs.GetRows()
	rs.Close
	
	Response.Write Ubound(arr,2)+1
	For i=0 To Ubound(arr,2)
		Response.Write SEPARATOR
		Response.Write arr(0,i) & SUBSEPARATOR & arr(1,i) & SUBSEPARATOR & arr(2,i)
		
		If arr(3,i)=0 OR arr(3,i)=1 Then
			Response.Write SUBSEPARATOR & arr(3,i) & SUBSEPARATOR & arr(4,i) &_
					SUBSEPARATOR & arr(5,i) & SUBSEPARATOR & arr(6,i) & SUBSEPARATOR & arr(7,i) &_
					SUBSEPARATOR & FormatDt(arr(8,i),"SQL_TM") & SUBSEPARATOR & arr(9,i)
		ElseIf arr(3,i)=2 Then
			Response.Write SUBSEPARATOR & arr(3,i) & SUBSEPARATOR & arr(4,i) &_
					SUBSEPARATOR & arr(5,i) & SUBSEPARATOR & arr(6,i) & SUBSEPARATOR & arr(7,i) &_
					SUBSEPARATOR & FormatDt(arr(8,i),"SQL_TM") & SUBSEPARATOR & arr(9,i) & SUBSEPARATOR & arr(10,i) &_
					SUBSEPARATOR & arr(11,i) & SUBSEPARATOR & CInt(arr(12,i)) & SUBSEPARATOR & arr(13,i) & SUBSEPARATOR & arr(14,i)
		End If
		Response.Write SUBSEPARATOR & arr(15,i) & SUBSEPARATOR & FormatDt(arr(16,i),"SQL_TM") & SUBSEPARATOR & arr(17,i)
	
		totalCount=totalCount+1
	Next
%>