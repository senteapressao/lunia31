<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : characterName
	' response : characterName,totalCount
	'			,{marketId,itemType,itemHash,stackedCount,itemSerial,instance,itemExpire,salePrice,expireDate,showLevel}
	'			,{marketId,itemType,itemHash,stackedCount,itemSerial,instance,itemExpire,petName,petLevel,isRare,rareProbability,enchantSerial,salePrice,expireDate,showLevel}

	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""

	' preparing request parameters
	Dim characterName
	If UBound(parameters)<0 Then Call Error("not enough parameter")
	characterName = Parameters(0)

	Dim totalCount : totalCount = 0

	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.Market_MyList"
		Call .InitCommand()
		Call .AppendParam("RETURN_VALUE",adInteger,adParamReturnValue,,null)
		Call .AppendParam("@characterName",adVarWChar,adParamInput,50,characterName)
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

	i=0
	Do Until rs.Eof
		If i>0 Then
			retString = retString & SEPARATOR
		End If
		
		retString = retString & rs(0)
		
		If rs(1)=0 OR rs(1)=1 Then
			retString = retString & SUBSEPARATOR & rs(1) & SUBSEPARATOR & rs(2) &_
					SUBSEPARATOR & rs(3) & SUBSEPARATOR & rs(4) & SUBSEPARATOR & rs(5) &_
					SUBSEPARATOR & FormatDt(rs(6),"SQL_TM")
		ElseIf rs(1)=2 Then
			retString = retString & SUBSEPARATOR & rs(1) & SUBSEPARATOR & rs(2) &_
					SUBSEPARATOR & rs(3) & SUBSEPARATOR & rs(4) & SUBSEPARATOR & rs(5) &_
					SUBSEPARATOR & FormatDt(rs(6),"SQL_TM") & SUBSEPARATOR & rs(7) & SUBSEPARATOR & rs(8) &_
					SUBSEPARATOR & CInt(rs(9)) & SUBSEPARATOR & rs(10) & SUBSEPARATOR & rs(11)
		End If
		retString = retString & SUBSEPARATOR & rs(12) & SUBSEPARATOR & FormatDt(rs(13),"SQL_TM") & SUBSEPARATOR & rs(14)

		totalCount=totalCount+1

		rs.MoveNext:i=i+1
	Loop
	rs.close

	retString = characterName & SEPARATOR & totalCount & SEPARATOR & retString

	Call OK(retString)
%>