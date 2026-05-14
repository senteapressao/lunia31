<%@ EnableSessionState=False %>
<%
Dim conn, connStr
connStr = "Provider=SQLOLEDB;Data Source=DESKTOP-S29QIQ5\SQLEXPRESS;Initial Catalog=v4_stage;user ID=SA;password=algoraSaturn2203@;"

Set conn = Server.CreateObject("ADODB.Connection")
conn.Open connStr

conn.Execute("DELETE FROM SquareStages WHERE serverName = 'Square Server'")
conn.Execute("DELETE FROM StageServers WHERE serverName = 'Square Server'")
conn.Execute("DELETE FROM Connections WHERE serverName = 'Square Server'")

Response.Write "Registros antigos do 'Square Server' removidos com sucesso!"

conn.Close
Set conn = Nothing
%>
