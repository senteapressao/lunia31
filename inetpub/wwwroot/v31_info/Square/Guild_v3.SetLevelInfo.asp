<!--#include file="../common.asp"-->
<!--#include file="../DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request : (POST)
	'			(subseparated: guildLevel,maxGuildExp,maxGuildPoint)
	' response : 0
	
	Dim i,j,k
	Dim sph,rs,retString,blsResult : blsResult = False
	
	Dim rowData
	rowData=getConvAscii(Request.BinaryRead(Request.TotalBytes))
	Parameters=Split(rowData,SEPARATOR)

	' preparing request parameters
	Dim arr
	
	For	i=0 To Ubound(Parameters)
		arr = Split(Parameters(i),SUBSEPARATOR)
		
		' call Stored procedure
		Set sph = new SPHelper
		with sph
			.DEBUG = False
			Set .cmd = Command
			.ConnStr = guildDBconnectionString
			.SPName = "dbo.public_SetLevelInfo"
			Call .InitCommand()
			Call .AppendParam("@guildLevel",adUnsignedTinyint,adParamInput,,arr(0))
			Call .AppendParam("@maxGuildExp",adInteger,adParamInput,,arr(1))
			Call .AppendParam("@maxGuildPoint",adInteger,adParamInput,,arr(2))
			Call .AppendParam("@maxGuildMember",adInteger,adParamInput,,arr(3))
			blsResult = .ExecRecordset()
		End with			
	Next

	Call Ok(0)
%>