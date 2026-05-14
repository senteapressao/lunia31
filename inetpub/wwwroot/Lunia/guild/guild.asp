<!--#include file="common.asp"-->
<!-- #include file="function.asp" -->
<!--#include file="DBhelper.asp"-->
<% 	
	 
Dim Uploader, File, FileSys, FilePath,ip, guildId, entrada,param, i, accountname, passwd, charactername, IpAddress, show, cnn, var, allow,sql,rs, guildName, message,una,shitty
una="<script>alert(""Not Authorized"");window.close();</script>"
shitty="Mozilla/4.0 (compatible; MSIE 7.0; Windows NT 6.2; WOW64; Trident/7.0; .NET4.0C; .NET4.0E; .NET CLR 2.0.50727; .NET CLR 3.0.30729; .NET CLR 3.5.30729; Tablet PC 2.0; .NET CLR 1.1.4322)"
if shitty < Request.ServerVariables("HTTP_USER_AGENT") then
	response.write(una)
	response.end
end if
entrada = FiltSQL(Request.QueryString)
ip = Request.ServerVariables ("remote_addr")
Set cnn = Server.CreateObject("ADODB.Connection")
cnn.open "PROVIDER=SQLOLEDB;DATA SOURCE=35.199.78.111;UID=sa;PWD=KjCa4F6gBZT%-RQAb78HsddJ34zRUvp3;DATABASE=v3_guild"
If cnn.errors.count = 0 or entrada <> "" Then   
   'Response.Write "Connected OK"
   allow=true
else
	allow=false
	response.write(una)
	response.end
End If
if  allow = false then 
	response.write(una)
	response.end
else
	entrada=FiltSQL(entrada)
	entrada=FiltHTML(entrada)
	entrada=FiltBR(entrada)
	entrada=FiltFlash(entrada)
	entrada = Decrypt(entrada)
	entrada=FiltSQL(entrada)
	entrada=FiltHTML(entrada)
	entrada=FiltBR(entrada)
	entrada=FiltFlash(entrada)
	entrada = Replace(entrada,"=","&")
	param=Split(entrada,"&")
	 'for i = 1 to uBound(param) 
	 'response.write(param(i)&"<BR/>")
	 'i= i+1
	 'next
	' recive the parameter from the first split
	if uBound(param)>7 then
		if (param(1)="") or (param(7)="") or (ip="") then
			response.write("Sem parametros de entrada!!")
		else 	
			accountname = param(1)
			passwd	=	param(3)
			charactername = param(7)
			IpAddress = Request.ServerVariables("remote_addr")
			session("accountname") = param(1)
			session("passwd") = param(3)
			session("charactername") = param(7)
			session("IpAddress") = ip
		end if
		else 
		allow = false
		response.write(una)
		response.end
		end if
if allow=true then
	' Getting guildId
	sql="SELECT guildid FROM [v3_guild].[dbo].[guild] where mastername='"&charactername&"';"
	Set rs=cnn.Execute(sql)
	guildId = rs.GetString()
	'response.write(guildId)
	' Geting guildName
	sql="SELECT guildName FROM [v3_guild].[dbo].[guild] where mastername='"&charactername&"';"
	Set rs=cnn.Execute(sql)
	guildName = rs.GetString()
	'response.write(guildName)
	' Geting guildName
	sql="SELECT message FROM [v3_guild].[dbo].[guild] where mastername='"&charactername&"';"
	Set rs=cnn.Execute(sql)
	message = rs.GetString()
	'response.write(message)
	rs.close
	cnn.close
else
	response.write(una)
	response.end
end if
end if
set rs = Nothing 
set cnn = Nothing 

%>

<html>
<meta http-equiv="Content-Type" content="text/html;charset=UTF-8" />
<head>
	<title>Lunia</title>
	<meta name="visual_author" content="Sage" name="logic_authour" name="Teles"> <!--All I ask is that you keep this here.-->
	<link rel="stylesheet" type="text/css" href="css/main.css" />
</head>
<body>
	<div class="container">
		<div class="heading">
			Minha Guilda
		</div>
		<div class="myInfo">
			<div class="iconArea">
				<img src="showicon.asp?guildid=<%=guildId%>" height="51" width="82">
			</div>
			<div class="personalArea">
				Você é o líder da <div class="guildName"><%=guildName%></div>
			</div>
		</div>
		<div class="blockquote">
			&ldquo;<%=message%>&rdquo;
		</div>
		
		<form id="formUpload" name="formUpload" enctype="multipart/form-data" method="POST" action="logic.asp" >
		<div class="fileArea">
			<input type="file" accept=".png" name="file1">
			<div class="instructions">
				O arquivo deve ser <strong>.png</strong><br/>e deve possuir menor de <strong>30kb</strong>.
			</div>
		</div>
		<div class="submitArea">
			<input class="submits" type="submit" name="submit" value=" ">
			<button class="cancel" onclick="window.close();"> </button>
		</div>
		</form>
	</div>
</body>
</html>
