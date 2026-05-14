<!--#include file="../common.asp"-->
<!--#include file="../DBhelper_NoTran.asp"-->
<!--METADATA TYPE="typelib" NAME="ADODB Type Library" UUID= "00000205-0000-0010-8000-00AA006D2EA4" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
	Call Init()
	' request  : floor
	' response : characterName
		
	Dim i,j,k
	Dim sphn,rs,blsResult : blsResult = False
	Dim retString : retString = ""	

	Dim stageGroupHash,accessLevel,floor
	If UBound(parameters)<2 Then Call Error("not enough parameter")
	stageGroupHash	= Parameters(0)
	accessLevel		= Parameters(1)
	floor			= Parameters(2)
	floor	= floor-1
	'response 변수 초기화
	
	' call Stored procedure	
	Set sphn = new SPHelper_NoTran
	with sphn
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = characterDBconnectionString
		.SPName = "dbo.TowerListLastPlayer"
		Call .InitCommand()
		Call .AppendParam("@stageGroupHash",adInteger,adParamInput,,stageGroupHash)
		Call .AppendParam("@accessLevel",adSmallInt,adParamInput,,accessLevel)
		Call .AppendParam("@floor",adTinyInt,adParamInput,,floor)
		blsResult = .ExecRecordset()
	End with
	
	If blsResult Then
		Set rs = sphn.rs
	End If
	set sphn = Nothing

	i=0
	Do Until rs.Eof
		If i>0 Then
			retString = retString & SEPARATOR
		End If
			
		retString = retString & rs(0)

		rs.MoveNext:i=i+1
	Loop
	rs.close
	
	Call OK(retString)
%>