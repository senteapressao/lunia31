<!--#include file="common.asp"-->
<!-- #include file="upLoadFunctions.asp" -->
<!--#include file="DBhelper.asp"-->
<!-- METADATA TYPE="typelib" FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->
<object runat="server" progid="ADODB.Command" id="Command" viewastext="viewastext"></object>
<%
' Create the FileUploader
Dim Uploader, File, FileSys, FilePath, param, continuar, mensagem, msg, size, atual, meio, ext, seuip
dim accountname, passwd, charactername, IpAddress, guildId
ext="png"
Set Uploader = New FileUploader
continuar = true
' This starts the upload process
Uploader.Upload()

' Check if any files were uploaded

If Uploader.Files.Count = 0 Then
    mensagem = "Nenhum arquivo enviado!"
Else
' Calling stored procedure to know the icon name and some security, of course :)
	accountname = session("accountname")
	passwd	=	session("passwd")
	charactername = session("charactername")
	IpAddress = session("IpAddress")
	' call Stored procedure	
	Dim sph,rs,blsResult : blsResult = False
	Set sph = new SPHelper
	with sph
		.DEBUG = False
		Set .cmd = Command
		.ConnStr = guildDBconnectionString
		.SPName = "dbo.check_UploadIcon"
		Call .InitCommand()
		Call .AppendParam("@accountname",adVarWChar,adParamInput,50,accountname)
		Call .AppendParam("@passwd",adVarWChar,adParamInput,80,passwd)
		Call .AppendParam("@charactername",adVarWChar,adParamInput,50,charactername)
		Call .AppendParam("@IpAddress",adVarWChar,adParamInput,50,IpAddress)
		Call .AppendParam("@guildId",adInteger,adParamOutput,,guildId)
		blsResult = .ExecNoRecords()
	End with

If blsResult=False Then
	if sph.frk_n4ErrorCode<>2 Then
		'Call Error(sph.frk_n4ErrorCode)
	Else
		'Call Error(sph.frk_strErrorText)
	End If
	continuar = false
End If
		guildId = sph.GetParamValue("@guildId")
		' Loop through the uploaded files
		
		For Each File In Uploader.Files.Items
		
		param = Split(File.FileName,".")
		size= int(File.FileSize/1024)
		
	if (param(UBound(param)) <> ext) or size > 30 or continuar = false Then
			continuar = false
		else
		
			File.FileName = guildId&"."&ext
			' Set upload Path and Filename to check if that file already exists
			FilePath = "C:\Inetpub\guildicon\"&File.FileName
		
			' Save the file
            File.SaveToDisk "C:\\Inetpub\\guildicon"
		end if 
		Next
		if continuar = true then
			mensagem = "Ícone trocado com sucesso, relogue praça para aplicar o efeito :)" 
		else
	
		mensagem = "Formato, tamanho ou autenticação inválidos!"
		msg = "Use ."&ext&" com no máximo <b>30KB</b>"
		meio = "A imagem atual possuí as seguintes  especificações :"
	
		atual = "Tamanho : "&size & " <b>KB</b> | Formato '<font size=""4""><b>."&param(UBound(param))&"</b></font>'"
		seuip = "Endereço IP : <font size=""4""><b>"&IpAddress&"</b></font>"
		end If
	end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=UTF-8" />
<style>
body {
    background: #fff;
}

.content {
    max-width: 500px;
    margin: auto;
    background: white;
    padding: 10px;
}
</style>
<title>Trocar Ícone da Guild</title>
</head>
<body >
<div style="color:#0000FF" align="center" class="content">
<font size="6"><b>
<center><p><%	response.write(mensagem)	%></p></center></b></font>
<center><p><%		response.write(msg)		%></p></center>
<center><p><%		response.write(meio)		%></p></center>
<center><p><%		response.write(atual)		%></p></center>
<center><p><%		response.write(seuip)		%></p></center>
</div>
</body>
</html>